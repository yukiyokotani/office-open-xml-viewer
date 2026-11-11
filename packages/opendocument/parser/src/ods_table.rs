//! One `table:table` (ODF Part 3 §9.1) as a worksheet: columns and rows with
//! their repeat counts (§19.675 / §19.680), cells and their values (§19.385
//! `office:value-type`), spans (§19.676), styles, drawings and annotations.
//!
//! Repeats are walked as runs, never expanded: a repeated row or cell is
//! materialized only when it carries content or a format that differs from
//! its column's default, so LibreOffice's trailing million-row filler costs
//! one run.

use super::{
    cells::XfBuilder, column_letters, date_serial, draw, duration_days, pt_to_px, Context,
    OdsSheet, MAX_COLS, MAX_ROWS, NS_TABLE,
};
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
/// The longest A1 cell reference (`XFD1048576`).
const MAX_CELL_REF_BYTES: usize = 10;

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
pub fn read_view_settings(root: roxmltree::Node<'_, '_>) -> HashMap<String, ViewSettings> {
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
        return tables;
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
    tables
}

/// One authored cell, copied to every cell its row and column repeats
/// cover. `cost` is what one copy retains (inline plus owned heap, from the
/// exhaustive `RetainedBytes` accounting), admitted before each copy.
#[derive(Clone)]
struct Template {
    cell: Rc<Cell>,
    cost: usize,
}

impl Template {
    fn new(cell: Cell) -> Self {
        let cost = super::retained_size(&cell);
        Self {
            cell: Rc::new(cell),
            cost,
        }
    }
}

/// A copy of `cell` at (`row`, `col`). The caller admitted its cost.
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
/// The default (xf, style name) of a column.
type ColumnDefault = dyn Fn(u32, &[Column]) -> (u32, Option<String>);

struct Column {
    first: u32,
    last: u32,
    px: u32,
    hidden: bool,
    level: u8,
    default_xf: u32,
    default_style: Option<String>,
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
    let default_xf_for_col = |col: u32, columns: &[Column]| -> (u32, Option<String>) {
        columns
            .iter()
            .find(|c| c.first <= col && col <= c.last)
            .map(|c| (c.default_xf, c.default_style.clone()))
            .unwrap_or((0, None))
    };

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
        next_row: 1,
        sheet_index: index,
    };
    walk_rows(
        ctx,
        table,
        0,
        false,
        &columns,
        &default_xf_for_col,
        &mut walker,
    )?;
    for shapes in table.children().filter(|n| xml::is(*n, NS_TABLE, "shapes")) {
        for node in shapes.children().filter(|n| n.is_element()) {
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
            ctx.budget
                .bytes(((*last - *first + 1) as usize).saturating_mul(24))?;
            for row in *first..=*last {
                row_heights.insert(row, px_to_pt(*px));
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
            ctx.budget
                .bytes(((*last - *first + 1) as usize).saturating_mul(48))?;
            for row in *first..=*last {
                if existing.insert(row) {
                    ctx.budget.cell(std::mem::size_of::<Row>())?;
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
    for c in &columns {
        for col in c.first..=c.last.min(c.first + 16_383) {
            if c.hidden {
                worksheet.col_hidden.insert(col, true);
            }
            if c.level > 0 {
                worksheet.col_outline_levels.insert(col, c.level);
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
    let font = ctx.xfs.default_font_model().clone();
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
    let geometry = draw::Grid::new(&column_px, &worksheet, default_row_px);
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
    // Per column: the blank xf seen so far (None: conflicting), and rows.
    let mut seen: Vec<(Option<u32>, u32, bool)> = vec![(None, 0, true); width + 1];
    let mut rows: Vec<roxmltree::Node<'_, '_>> = Vec::new();
    collect_row_nodes(table, &mut rows);
    // Only a table that reaches the last row can format whole columns, and
    // only columns whose last row paints differently from the column
    // default are candidates; the scan below visits those columns only.
    let mut total = 0u64;
    for row in &rows {
        total = total.saturating_add(u64::from(repeat(*row, "number-rows-repeated")?));
    }
    if total < u64::from(MAX_ROWS) {
        return Ok(());
    }
    let mut candidates: Vec<(usize, usize)> = Vec::new();
    if let Some(last_row) = rows.last() {
        let row_default = last_row.attribute((NS_TABLE, "default-cell-style-name"));
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
            let own = cell.attribute((NS_TABLE, "style-name")).or(row_default);
            let mut c = col;
            while c <= last {
                let column = columns
                    .iter()
                    .find(|k| k.first as usize <= c && c <= k.last as usize);
                let segment_last = column.map_or(last, |k| (k.last as usize).min(last));
                let style = own
                    .map(str::to_string)
                    .or_else(|| column.and_then(|k| k.default_style.clone()));
                let xf = ctx
                    .xfs
                    .xf_for(&ctx.store, &mut ctx.data, &ctx.faces, style.as_deref());
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
    let mut next_row = 1u32;
    for row in rows {
        if next_row > MAX_ROWS {
            break;
        }
        let count = repeat(row, "number-rows-repeated")?.min(MAX_ROWS - next_row + 1);
        next_row += count;
        let row_default = row.attribute((NS_TABLE, "default-cell-style-name"));
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
            let own = cell.attribute((NS_TABLE, "style-name")).or(row_default);
            let mut c = col;
            while c <= last {
                let column = columns
                    .iter()
                    .find(|k| k.first as usize <= c && c <= k.last as usize);
                let segment_last = column.map_or(last, |k| (k.last as usize).min(last));
                let style = own
                    .map(str::to_string)
                    .or_else(|| column.and_then(|k| k.default_style.clone()));
                let xf = ctx
                    .xfs
                    .xf_for(&ctx.store, &mut ctx.data, &ctx.faces, style.as_deref());
                let slots = candidates
                    .iter()
                    .filter(|(a, b)| *a <= segment_last && c <= *b)
                    .flat_map(|(a, b)| (*a).max(c)..=(*b).min(segment_last));
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
        for c in column.first..=column.last + 1 {
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
                    default_style: column.default_style.clone(),
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
) {
    for node in parent.children().filter(|n| n.is_element()) {
        if node.tag_name().namespace() != Some(NS_TABLE) {
            continue;
        }
        match node.tag_name().name() {
            "table-row-group" | "table-header-rows" | "table-rows" => collect_row_nodes(node, out),
            "table-row" => out.push(node),
            _ => {}
        }
    }
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
                let default_style = node
                    .attribute((NS_TABLE, "default-cell-style-name"))
                    .map(str::to_string);
                let default_xf = ctx.xfs.xf_for(
                    &ctx.store,
                    &mut ctx.data,
                    &ctx.faces,
                    default_style.as_deref(),
                );
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
    style_maps: Vec<(String, xlsx_model::CellRange)>,
    /// (style, first column, last column) → its last range, to extend it
    /// down over the next rows.
    style_map_last: HashMap<(String, u32, u32), usize>,
    next_row: u32,
    #[allow(dead_code)]
    sheet_index: usize,
}

#[allow(clippy::too_many_arguments)]
fn walk_rows<'a, 'input>(
    ctx: &mut Context,
    parent: roxmltree::Node<'a, 'input>,
    level: u8,
    group_hidden: bool,
    columns: &[Column],
    column_default: &ColumnDefault,
    walker: &mut RowWalker<'a, 'input>,
) -> Result<(), String> {
    for node in parent.children().filter(|n| n.is_element()) {
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
                    column_default,
                    walker,
                )?;
            }
            "table-header-rows" | "table-rows" => walk_rows(
                ctx,
                node,
                level,
                group_hidden,
                columns,
                column_default,
                walker,
            )?,
            "table-row" => read_row(
                ctx,
                node,
                level,
                group_hidden,
                columns,
                column_default,
                walker,
            )?,
            _ => {}
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn read_row<'a, 'input>(
    ctx: &mut Context,
    node: roxmltree::Node<'a, 'input>,
    level: u8,
    group_hidden: bool,
    columns: &[Column],
    column_default: &ColumnDefault,
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
    let row_default = node.attribute((NS_TABLE, "default-cell-style-name"));

    // The cells of one occurrence of this row.
    let mut templates: Vec<(u32, Template)> = Vec::new();
    let mut col = 1u32;
    for cell in node.children().filter(|n| n.is_element()) {
        let covered = xml::is(cell, NS_TABLE, "covered-table-cell");
        if !covered && !xml::is(cell, NS_TABLE, "table-cell") {
            continue;
        }
        if col > MAX_COLS {
            break;
        }
        let repeat_count = repeat(cell, "number-columns-repeated")?;
        let cell_last = col.saturating_add(repeat_count - 1).min(MAX_COLS);
        let own_style = cell.attribute((NS_TABLE, "style-name"));
        let (value, formula) = cell_value(
            ctx,
            cell,
            own_style,
            row_default,
            col,
            columns,
            column_default,
        )?;
        let has_value = !matches!(value, CellValue::Empty);
        // The prototype of the last format seen in this authored cell; its
        // copies are made from it.
        let mut prototype: Option<Template> = None;
        // Drawings, annotations, links and spans belong to the first row of
        // a repeated block only when it is not repeated (LibreOffice never
        // repeats rows that carry them).
        // Walk the repeated cells as segments of constant column default.
        let mut c = col;
        while c <= cell_last {
            let segment_last = columns
                .iter()
                .find(|column| column.first <= c && c <= column.last)
                .map_or(cell_last, |column| column.last.min(cell_last));
            let (column_xf, column_style) = column_default(c, columns);
            let style = own_style
                .or(row_default)
                .map(str::to_string)
                .or(column_style);
            let style_name = style.as_deref().unwrap_or("Default");
            if ctx
                .store
                .find("table-cell", style_name)
                .is_some_and(|s| !s.maps.is_empty())
            {
                record_style_map(ctx, walker, style_name, (first, c), (last, segment_last))?;
            }
            let xf = ctx
                .xfs
                .xf_for(&ctx.store, &mut ctx.data, &ctx.faces, style.as_deref());
            let xf = adjust_xf_for_value(&mut ctx.xfs, xf, cell);
            // An empty cell is materialized only when it paints differently
            // from its column's default (fill or borders); its font, number
            // format and alignment have nothing to show.
            let paints = ctx.xfs.blank_signature(xf) != ctx.xfs.blank_signature(column_xf);
            if has_value || paints || formula.is_some() {
                let template = match prototype.take() {
                    Some(t) if t.cell.style_index == Some(xf) => t,
                    _ => Template::new(Cell {
                        col,
                        row: first,
                        value: copy_value(&value),
                        style_index: Some(xf),
                        formula: formula.clone(),
                        show_phonetic: false,
                    }),
                };
                for c in c..=segment_last {
                    // Admitted before the copy is made.
                    ctx.budget.cell(template.cost)?;
                    templates.push((c, template.clone()));
                }
                prototype = Some(template);
            }
            c = segment_last + 1;
        }
        for c in col..=cell_last {
            if covered {
                break;
            }
            let span_cols = repeat(cell, "number-columns-spanned")?;
            let span_rows = repeat(cell, "number-rows-spanned")?;
            if span_cols > 1 || span_rows > 1 {
                ctx.budget
                    .bytes(((last - first + 1) as usize).saturating_mul(32))?;
                for row in first..=last {
                    walker.merges.push(MergeCell {
                        top: row,
                        left: c,
                        bottom: row.saturating_add(span_rows - 1).min(MAX_ROWS),
                        right: c.saturating_add(span_cols - 1).min(MAX_COLS),
                    });
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
        // the block carries them (§19.675 / §19.680), each admitted to the
        // budget before it is stored. Drawings are anchored once, at the
        // block's first cell, as LibreOffice anchors them.
        if !covered || cell.children().any(|n| n.is_element()) {
            let link = cell
                .descendants()
                .find(|n| xml::is(*n, NS_TEXT, "a"))
                .and_then(|a| a.attribute((NS_XLINK, "href")));
            let annotation = cell
                .children()
                .find(|n| xml::is(*n, NS_OFFICE, "annotation"));
            if link.is_some() || annotation.is_some() {
                // One prototype each, built once (a note's text with its
                // `text:s` spaces expanded); every copy is admitted at the
                // prototype's exact retained size before it is made. A
                // note's cell reference is also kept in `comment_refs`.
                let link = link.map(|href| hyperlink(col, first, href));
                let link_cost = link.as_ref().map_or(0, super::retained_size);
                let note = annotation.map(|a| comment(a, super::paragraph_text(a), first, col));
                let note_cost = note.as_ref().map_or(0, |n| {
                    super::retained_size(n)
                        .saturating_add(std::mem::size_of::<String>())
                        .saturating_add(2 * MAX_CELL_REF_BYTES)
                });
                for row in first..=last {
                    for c in col..=cell_last {
                        if let Some(link) = &link {
                            ctx.budget.bytes(link_cost)?;
                            walker.hyperlinks.push(Hyperlink {
                                col: c,
                                row,
                                ..link.clone()
                            });
                        }
                        if let Some(note) = &note {
                            ctx.budget.bytes(note_cost)?;
                            walker.comments.push(xlsx_model::XlsxComment {
                                cell_ref: format!("{}{}", column_letters(c), row),
                                ..note.clone()
                            });
                        }
                    }
                }
            }
            for child in cell.children().filter(|n| is_drawing(*n)) {
                walker.anchored.push(Anchored {
                    node: child,
                    cell: Some((first - 1, col - 1)),
                });
            }
        }
        col = cell_last + 1;
    }
    if !templates.is_empty() {
        for row in first..=last {
            // The first occurrence was admitted with its templates; each
            // repeat is admitted before it is copied.
            if row != first {
                for (_, t) in &templates {
                    ctx.budget.cell(t.cost)?;
                }
            }
            ctx.budget.bytes(std::mem::size_of::<Row>())?;
            let cells: Vec<Cell> = templates
                .iter()
                .map(|(c, t)| copy_cell(&t.cell, row, *c))
                .collect();
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
    ctx: &mut Context,
    walker: &mut RowWalker<'_, '_>,
    style: &str,
    (top, left): (u32, u32),
    (bottom, right): (u32, u32),
) -> Result<(), String> {
    let key = (style.to_string(), left, right);
    if let Some(index) = walker.style_map_last.get(&key) {
        let range = &mut walker.style_maps[*index].1;
        if range.bottom.saturating_add(1) == top {
            range.bottom = bottom;
            return Ok(());
        }
    }
    // The range and its index key each hold the style name.
    ctx.budget
        .bytes(64usize.saturating_add(style.len().saturating_mul(2)))?;
    walker.style_map_last.insert(key, walker.style_maps.len());
    walker.style_maps.push((
        style.to_string(),
        xlsx_model::CellRange {
            top,
            left,
            bottom,
            right,
        },
    ));
    Ok(())
}

/// A cell whose value type has no matching data style is shown as
/// LibreOffice shows that type (§19.385): dates and times with the default
/// ISO-like date / time formats, percentages with `0%` decimals of the value.
fn adjust_xf_for_value(xfs: &mut XfBuilder, xf: u32, cell: roxmltree::Node<'_, '_>) -> u32 {
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
        _ => xf,
    }
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

/// The value of a cell (§19.385 value types) and its formula.
fn cell_value(
    ctx: &mut Context,
    cell: roxmltree::Node<'_, '_>,
    own_style: Option<&str>,
    row_default: Option<&str>,
    col: u32,
    columns: &[Column],
    column_default: &ColumnDefault,
) -> Result<(CellValue, Option<String>), String> {
    let formula = cell
        .attribute((NS_TABLE, "formula"))
        .map(super::extras::convert_formula);
    let value_type = cell.attribute((NS_OFFICE, "value-type"));
    let calc_type = cell.attribute((NS_CALCEXT, "value-type"));
    let number = |name: &str| {
        cell.attribute((NS_OFFICE, name))
            .and_then(|v| v.trim().parse::<f64>().ok())
    };
    if calc_type == Some("error") {
        let error = super::paragraph_text(cell);
        return Ok((CellValue::Error { error }, formula));
    }
    let value = match value_type {
        Some("float") | Some("percentage") | Some("currency") => match number("value") {
            Some(n) => CellValue::Number { number: n },
            None => CellValue::Empty,
        },
        Some("date") => match cell
            .attribute((NS_OFFICE, "date-value"))
            .and_then(|v| date_serial(v, ctx.null_date_shift, ctx.date1904))
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
                let (column_xf, column_style) = column_default(col, columns);
                let _ = column_xf;
                let style = own_style
                    .or(row_default)
                    .map(str::to_string)
                    .or(column_style);
                let xf = ctx
                    .xfs
                    .xf_for(&ctx.store, &mut ctx.data, &ctx.faces, style.as_deref());
                let (text, runs) = rich_text(ctx, cell, xf)?;
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
    // LibreOffice stores a formula's empty-string result as the number 0
    // with an empty displayed paragraph; it shows nothing.
    let value = match value {
        CellValue::Number { number }
            if number == 0.0
                && formula.is_some()
                && cell.children().any(|p| xml::is(p, NS_TEXT, "p"))
                && super::paragraph_text(cell).is_empty() =>
        {
            text_value(String::new(), None)
        }
        other => other,
    };
    Ok((value, formula))
}

fn text_value(text: String, runs: Option<Vec<Run>>) -> CellValue {
    CellValue::Text {
        text,
        runs,
        phonetic_runs: Vec::new(),
        phonetic_pr: None,
    }
}

/// Cell text with `text:span` runs (§6.1.7) when any span formats text.
/// Runs borrow their style names from the document; each run's font is a
/// copy (its family name included), admitted as it is made.
fn rich_text(
    ctx: &mut Context,
    cell: roxmltree::Node<'_, '_>,
    xf: u32,
) -> Result<(String, Option<Vec<Run>>), String> {
    let mut runs: Vec<(String, Option<&str>)> = Vec::new();
    let mut first = true;
    for p in cell
        .children()
        .filter(|n| xml::is(*n, NS_TEXT, "p") || xml::is(*n, NS_TEXT, "h"))
    {
        if !first {
            push_run(&mut runs, "\n", None);
        }
        first = false;
        collect_runs(p, None, &mut runs);
    }
    let text: String = runs.iter().map(|(t, _)| t.as_str()).collect();
    if runs.iter().all(|(_, style)| style.is_none()) {
        return Ok((text, None));
    }
    let cell_font = ctx.xfs.font_of_xf(xf);
    let base = xlsx_model::RunFont {
        bold: cell_font.bold,
        italic: cell_font.italic,
        underline: cell_font.underline,
        strike: cell_font.strike,
        size: Some(cell_font.size),
        color: cell_font.color.clone(),
        name: cell_font.name.clone(),
        underline_style: cell_font.underline_style.clone(),
        vert_align: cell_font.vert_align.clone(),
        ..xlsx_model::RunFont::default()
    };
    let mut out = Vec::with_capacity(runs.len());
    for (text, style) in runs {
        let run = Run {
            text,
            font: Some(match style {
                Some(style) => ctx.xfs.run_font(&ctx.store, &ctx.faces, &cell_font, style),
                None => base.clone(),
            }),
        };
        ctx.budget.bytes(super::retained_size(&run))?;
        out.push(run);
    }
    Ok((text, Some(out)))
}

fn push_run<'a>(runs: &mut Vec<(String, Option<&'a str>)>, text: &str, style: Option<&'a str>) {
    if let Some(last) = runs.last_mut() {
        if last.1 == style {
            last.0.push_str(text);
            return;
        }
    }
    runs.push((text.to_string(), style));
}

fn collect_runs<'a>(
    node: roxmltree::Node<'a, '_>,
    style: Option<&'a str>,
    runs: &mut Vec<(String, Option<&'a str>)>,
) {
    for child in node.children() {
        if child.is_text() {
            push_run(runs, child.text().unwrap_or(""), style);
            continue;
        }
        if child.tag_name().namespace() != Some(NS_TEXT) {
            continue;
        }
        match child.tag_name().name() {
            // `content.xml` is bounded by `check_space_runs` before reading.
            "s" => push_run(
                runs,
                &" ".repeat(crate::odf_text::space_count(child)),
                style,
            ),
            "tab" => push_run(runs, "\t", style),
            "line-break" => push_run(runs, "\n", style),
            "span" => {
                let inner = child.attribute((NS_TEXT, "style-name")).or(style);
                collect_runs(child, inner, runs);
            }
            "note" | "ruby-text" => {}
            _ => collect_runs(child, style, runs),
        }
    }
}
