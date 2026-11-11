//! Presentation tables (ODF 1.3 Part 3 §9.1 `table:table` in a
//! `draw:frame`, §16.17 table templates) mapped onto the PPTX table model.

use crate::odp_geometry::Placed;
use crate::odp_media::Media;
use crate::odp_paint::{self as paint, PaintContext};
use crate::odp_style::{
    Nd, Props, Scope, Sheet, LOEXT_GRAPHIC, PARAGRAPH, TABLE_CELL, TABLE_COLUMN, TABLE_ROW,
};
use crate::odp_text::{self as text, ObjectStyles, TextEnv, TextPayload};
use crate::odp_units::{color, length_emu};
use crate::xml::{self, NS_FO, NS_STYLE, NS_TABLE};
use pptx_model::{SlideElement, Stroke, TableCell, TableElement, TableRow};
use std::collections::HashMap;

/// Implementation bounds on a table's grid.
const MAX_COLUMNS: usize = 1024;
const MAX_ROWS: usize = 16_384;
const MAX_CELLS: usize = 262_144;

#[allow(clippy::too_many_arguments)]
pub fn table<'a>(
    sheet: &'a Sheet<'a>,
    media: &mut Media<'_>,
    node: Nd<'a>,
    scope: Scope,
    base: &'a str,
    placed: &Placed,
    frame_styles: &ObjectStyles<'a>,
    slide_index: usize,
    slide_count: usize,
    pages: &HashMap<String, usize>,
    text_work: &std::cell::Cell<usize>,
    out: &mut Vec<SlideElement>,
) -> Result<(), String> {
    let template = node
        .attribute((NS_TABLE, "template-name"))
        .and_then(|name| template(sheet, name));
    let flag = |name: &str| node.attribute((NS_TABLE, name)) == Some("true");
    let mut cols: Vec<i64> = Vec::new();
    let mut rows_nodes: Vec<Nd<'a>> = Vec::new();
    for child in node.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "table-column" => push_columns(sheet, scope, child, &mut cols)?,
            "table-columns" | "table-header-columns" => {
                for c in child
                    .children()
                    .filter(|n| xml::is(*n, NS_TABLE, "table-column"))
                {
                    push_columns(sheet, scope, c, &mut cols)?;
                }
            }
            "table-row" => push_rows(child, &mut rows_nodes)?,
            "table-rows" | "table-header-rows" => {
                for r in child
                    .children()
                    .filter(|n| xml::is(*n, NS_TABLE, "table-row"))
                {
                    push_rows(r, &mut rows_nodes)?;
                }
            }
            _ => {}
        }
    }
    let row_count = rows_nodes.len();
    let mut rows = Vec::with_capacity(row_count);
    let mut cells_total = 0usize;
    // One text payload for the whole table: the cells' text as it is read,
    // and each cell record (borders, fill, strings) before it is kept.
    let payload = TextPayload::new(text_work);
    for (r, row) in rows_nodes.iter().enumerate() {
        let row_props = row
            .attribute((NS_TABLE, "style-name"))
            .map(|n| sheet.chain(scope, "table-row", n))
            .transpose()?
            .unwrap_or_default();
        let mut rp = Props::new();
        rp.extend_styles(&row_props, TABLE_ROW);
        let height = rp
            .get(NS_STYLE, "row-height")
            .or_else(|| rp.get(NS_STYLE, "min-row-height"))
            .and_then(length_emu)
            .unwrap_or(0.0);
        let default_cell = row.attribute((NS_TABLE, "default-cell-style-name"));
        let mut cells = Vec::new();
        let mut c = 0usize;
        for cell in row.children().filter(|n| n.is_element()) {
            let covered = xml::is(cell, NS_TABLE, "covered-table-cell");
            if !(covered || xml::is(cell, NS_TABLE, "table-cell")) {
                continue;
            }
            let repeat = cell
                .attribute((NS_TABLE, "number-columns-repeated"))
                .and_then(|v| v.parse::<usize>().ok())
                .unwrap_or(1)
                .clamp(1, MAX_COLUMNS);
            for _ in 0..repeat {
                cells_total += 1;
                if cells_total > MAX_CELLS || c >= MAX_COLUMNS {
                    return Err("OpenDocument table exceeds the cell limit".to_string());
                }
                let template_style = template
                    .as_ref()
                    .and_then(|t| t.style_for(r, c, row_count, cols.len(), &flag));
                let style_name = cell.attribute((NS_TABLE, "style-name")).or(default_cell);
                let mut chain = style_name
                    .map(|n| sheet.chain(scope, "table-cell", n))
                    .transpose()?
                    .unwrap_or_default();
                if let Some(t) = template_style.and_then(|n| sheet.find(scope, "table-cell", n)) {
                    chain.push(t);
                }
                let mut built = build_cell(
                    sheet,
                    media,
                    cell,
                    covered,
                    &chain,
                    scope,
                    base,
                    frame_styles,
                    slide_index,
                    slide_count,
                    pages,
                    &payload,
                )?;
                let body = built.text_body.take();
                payload.charge(text::record_bytes(&built)?)?;
                built.text_body = body;
                cells.push(built);
                c += 1;
            }
        }
        rows.push(TableRow {
            height: height.round() as i64,
            cells,
        });
    }
    // Merge flags from spans (§19.680 number-columns-spanned / rows-spanned).
    let spans: Vec<Vec<(u32, u32)>> = rows_nodes
        .iter()
        .map(|row| {
            row.children()
                .filter(|n| {
                    xml::is(*n, NS_TABLE, "table-cell")
                        || xml::is(*n, NS_TABLE, "covered-table-cell")
                })
                .flat_map(|cell| {
                    let repeat = cell
                        .attribute((NS_TABLE, "number-columns-repeated"))
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1)
                        .clamp(1, MAX_COLUMNS);
                    let span = |name: &str| {
                        cell.attribute((NS_TABLE, name))
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1)
                            .max(1)
                    };
                    std::iter::repeat_n(
                        (span("number-columns-spanned"), span("number-rows-spanned")),
                        repeat,
                    )
                })
                .collect()
        })
        .collect();
    // Spans are clamped to the grid that exists (a span past the last row
    // or column covers what is there), so the model never claims more cells
    // than the table has and the merge walk below stays inside the grid.
    let row_total = rows.len();
    for (r, row) in rows.iter_mut().enumerate() {
        let row_len = row.cells.len();
        for (c, cell) in row.cells.iter_mut().enumerate() {
            let (cs, rs) = spans[r].get(c).copied().unwrap_or((1, 1));
            cell.grid_span = clamp_span(cs, row_len - c);
            cell.row_span = clamp_span(rs, row_total - r);
        }
    }
    // Every step of the walk marks an existing cell, and each covered cell
    // is marked once in a well-formed table (spans do not overlap), so the
    // walk is bounded by the cell count; overlapping spans that would mark
    // more cells than the table has are malformed.
    let mut visits = 0usize;
    for r in 0..rows.len() {
        for c in 0..rows[r].cells.len() {
            let (cs, rs) = (
                rows[r].cells[c].grid_span as usize,
                rows[r].cells[c].row_span as usize,
            );
            if rows[r].cells[c].h_merge || rows[r].cells[c].v_merge {
                continue;
            }
            for dr in 0..rs {
                // A row too short to hold the span's columns ends it.
                let Some(row) = rows.get_mut(r + dr) else {
                    break;
                };
                if row.cells.len() <= c {
                    break;
                }
                let end = c.saturating_add(cs).min(row.cells.len());
                for (dc, target) in row.cells[c..end].iter_mut().enumerate() {
                    if dr == 0 && dc == 0 {
                        continue;
                    }
                    visits += 1;
                    if visits > cells_total {
                        return Err("OpenDocument table spans overlap".to_string());
                    }
                    if dr == 0 {
                        target.h_merge = true;
                    } else {
                        target.v_merge = true;
                    }
                }
            }
        }
    }
    // Columns without widths share the frame width; rows without heights
    // share the remaining height.
    let ncols = rows
        .iter()
        .map(|r| r.cells.len())
        .max()
        .unwrap_or(0)
        .max(cols.len());
    while cols.len() < ncols {
        cols.push(0);
    }
    // File lengths may be huge: sums and differences saturate.
    let known: i64 = cols.iter().fold(0i64, |sum, w| sum.saturating_add(*w));
    let unknown = cols.iter().filter(|w| **w == 0).count() as i64;
    if unknown > 0 {
        let share = ((placed.w.round() as i64).saturating_sub(known).max(0)) / unknown;
        for w in cols.iter_mut().filter(|w| **w == 0) {
            *w = share;
        }
    }
    let known_h: i64 = rows
        .iter()
        .fold(0i64, |sum, r| sum.saturating_add(r.height));
    let unknown_h = rows.iter().filter(|r| r.height == 0).count() as i64;
    if unknown_h > 0 {
        let share = ((placed.h.round() as i64).saturating_sub(known_h).max(0)) / unknown_h;
        for r in rows.iter_mut().filter(|r| r.height == 0) {
            r.height = share;
        }
    }
    let table = TableElement {
        math_fallback: None,
        background: None,
        id: None,
        x: placed.x.round() as i64,
        y: placed.y.round() as i64,
        width: placed.w.round() as i64,
        height: placed.h.round() as i64,
        rotation: placed.rot,
        flip_h: false,
        flip_v: false,
        cols,
        rows,
        rtl: false,
    };
    out.push(SlideElement::Table(table));
    Ok(())
}

/// A file span clamped to the `available` cells from its origin (at least 1).
fn clamp_span(span: u32, available: usize) -> u32 {
    let available = u32::try_from(available.max(1)).unwrap_or(u32::MAX);
    span.clamp(1, available)
}

fn push_columns<'a>(
    sheet: &Sheet<'a>,
    scope: Scope,
    column: Nd<'a>,
    cols: &mut Vec<i64>,
) -> Result<(), String> {
    let repeat = column
        .attribute((NS_TABLE, "number-columns-repeated"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(1)
        .clamp(1, MAX_COLUMNS);
    let chain = column
        .attribute((NS_TABLE, "style-name"))
        .map(|n| sheet.chain(scope, "table-column", n))
        .transpose()?
        .unwrap_or_default();
    let mut props = Props::new();
    props.extend_styles(&chain, TABLE_COLUMN);
    let width = props
        .get(NS_STYLE, "column-width")
        .and_then(length_emu)
        .unwrap_or(0.0)
        .round() as i64;
    for _ in 0..repeat {
        if cols.len() >= MAX_COLUMNS {
            return Err("OpenDocument table exceeds the column limit".to_string());
        }
        cols.push(width);
    }
    Ok(())
}

fn push_rows<'a>(row: Nd<'a>, rows: &mut Vec<Nd<'a>>) -> Result<(), String> {
    let repeat = row
        .attribute((NS_TABLE, "number-rows-repeated"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(1)
        .clamp(1, MAX_ROWS);
    for _ in 0..repeat {
        if rows.len() >= MAX_ROWS {
            return Err("OpenDocument table exceeds the row limit".to_string());
        }
        rows.push(row);
    }
    Ok(())
}

/// A table template (§16.17 `table:table-template`): cell style names per
/// table region.
struct Template<'a> {
    regions: HashMap<&'a str, &'a str>,
}

fn template<'a>(sheet: &Sheet<'a>, name: &str) -> Option<Template<'a>> {
    let node = sheet.named("table-template", name)?;
    let regions = node
        .children()
        .filter(|n| n.is_element())
        .filter_map(|n| Some((n.tag_name().name(), n.attribute((NS_TABLE, "style-name"))?)))
        .collect();
    Some(Template { regions })
}

impl<'a> Template<'a> {
    fn style_for(
        &self,
        r: usize,
        c: usize,
        rows: usize,
        cols: usize,
        flag: &dyn Fn(&str) -> bool,
    ) -> Option<&'a str> {
        let get = |k: &str| self.regions.get(k).copied();
        if r == 0 && flag("use-first-row-styles") {
            if let Some(s) = get("first-row") {
                return Some(s);
            }
        }
        if r + 1 == rows && flag("use-last-row-styles") {
            if let Some(s) = get("last-row") {
                return Some(s);
            }
        }
        if c == 0 && flag("use-first-column-styles") {
            if let Some(s) = get("first-column") {
                return Some(s);
            }
        }
        if c + 1 == cols && flag("use-last-column-styles") {
            if let Some(s) = get("last-column") {
                return Some(s);
            }
        }
        if flag("use-banding-rows-styles") && r % 2 == 1 {
            if let Some(s) = get("odd-rows") {
                return Some(s);
            }
        }
        if flag("use-banding-columns-styles") && c % 2 == 1 {
            if let Some(s) = get("odd-columns") {
                return Some(s);
            }
        }
        get("body")
    }
}

/// `fo:border*` (§20.183) "width style colour" as a line.
fn border(value: Option<&str>) -> Option<Stroke> {
    let value = value?;
    if value == "none" || value.contains("hidden") {
        return None;
    }
    let mut width = None;
    let mut colour = None;
    let mut style = "solid";
    for part in value.split_whitespace() {
        if let Some(c) = color(part) {
            colour = Some(c);
        } else if let Some(w) = length_emu(part) {
            width = Some(w);
        } else {
            style = match part {
                "thin" => {
                    width = Some(0.75 * 12_700.0);
                    style
                }
                "medium" => {
                    width = Some(12_700.0 * 1.5);
                    style
                }
                "thick" => {
                    width = Some(12_700.0 * 3.0);
                    style
                }
                other => other,
            };
        }
    }
    if style == "none" || style == "hidden" {
        return None;
    }
    let width = width.unwrap_or(12_700.0);
    if width <= 0.0 {
        return None;
    }
    Some(Stroke {
        color: colour.unwrap_or_else(|| "000000".into()),
        width: width.round() as i64,
        fill: None,
        dash_style: match style {
            "dotted" => Some("sysDot".into()),
            "dashed" => Some("dash".into()),
            _ => None,
        },
        custom_dash: Vec::new(),
        line_cap: None,
        line_join: None,
        miter_limit: None,
        alignment: None,
        head_end: None,
        tail_end: None,
        cmpd: (style == "double").then(|| "dbl".into()),
    })
}

#[allow(clippy::too_many_arguments)]
fn build_cell<'a>(
    sheet: &'a Sheet<'a>,
    media: &mut Media<'_>,
    cell: Nd<'a>,
    covered: bool,
    chain: &[Nd<'a>],
    scope: Scope,
    base: &'a str,
    frame_styles: &ObjectStyles<'a>,
    slide_index: usize,
    slide_count: usize,
    pages: &HashMap<String, usize>,
    payload: &TextPayload<'_>,
) -> Result<TableCell, String> {
    let mut graphic = Props::new();
    graphic.extend_styles(chain, LOEXT_GRAPHIC);
    graphic.extend_styles(chain, crate::odp_style::GRAPHIC);
    let mut cell_props = Props::new();
    cell_props.extend_styles(chain, TABLE_CELL);
    let mut para = Props::new();
    para.extend_styles(chain, PARAGRAPH);
    let side = |name: &str| {
        cell_props
            .get(NS_FO, name)
            .or_else(|| para.get(NS_FO, name))
            .or_else(|| cell_props.get(NS_FO, "border"))
            .or_else(|| para.get(NS_FO, "border"))
    };
    let mut paint_ctx = PaintContext {
        sheet,
        media,
        base,
        slide_background: None,
    };
    let fill = if paint::has_fill_attribute(&graphic) {
        paint::fill(&graphic, &mut paint_ctx, (0.0, 0.0))
    } else {
        paint::background_color(&cell_props).map(|c| pptx_model::Fill::Solid { color: c })
    };
    let background = match &fill {
        Some(pptx_model::Fill::Solid { color }) => Some(color[..6.min(color.len())].to_string()),
        _ => None,
    };
    let mut object = frame_styles.clone();
    let mut cell_chain = chain.to_vec();
    cell_chain.extend(object.chain.iter().copied());
    object.chain = cell_chain;
    object.outline_master = None;
    object.automatic = None;
    let env = TextEnv {
        sheet,
        scope,
        base,
        object,
        slide_number: slide_index + 1,
        slide_count,
        footer: None,
        header: None,
        date_time: None,
        background,
        pages,
        payload,
    };
    let mut body = if covered {
        None
    } else {
        text::text_body(cell, &env, media, &graphic)?
    };
    if let Some(body) = &mut body {
        if let Some(anchor) = cell_props.get(NS_STYLE, "vertical-align") {
            body.vertical_anchor = match anchor {
                "middle" => "ctr",
                "bottom" => "b",
                _ => "t",
            }
            .to_string();
        }
        if cell_props
            .get(NS_STYLE, "rotation-angle")
            .and_then(|a| a.parse::<f64>().ok())
            == Some(90.0)
        {
            body.vert = "vert270".into();
        } else if cell_props
            .get(NS_STYLE, "rotation-angle")
            .and_then(|a| a.parse::<f64>().ok())
            == Some(270.0)
        {
            body.vert = "vert".into();
        }
        // Cell text frame (§20.230 table-cell-properties): wrapping, and
        // shrink-to-fit (§20.360) as for a shape's text.
        body.wrap = if cell_props.get(NS_FO, "wrap-option") == Some("no-wrap") {
            "none"
        } else {
            "square"
        }
        .into();
        if cell_props.get(NS_STYLE, "shrink-to-fit") == Some("true") {
            body.auto_fit = "norm".into();
            body.shrink_on_overflow = true;
        }
        // Cell padding (§20.230 fo:padding in table-cell-properties) over the
        // graphic padding the text frame read.
        let padding = |side: &str| {
            cell_props
                .get(NS_FO, side)
                .or_else(|| cell_props.get(NS_FO, "padding"))
                .and_then(length_emu)
                .map(|v| v.round() as i64)
        };
        if let Some(v) = padding("padding-left") {
            body.l_ins = v;
        }
        if let Some(v) = padding("padding-right") {
            body.r_ins = v;
        }
        if let Some(v) = padding("padding-top") {
            body.t_ins = v;
        }
        if let Some(v) = padding("padding-bottom") {
            body.b_ins = v;
        }
    }
    Ok(TableCell {
        text_body: body,
        horz_overflow: None,
        fill,
        has_direct_fill: true,
        text_color: None,
        border_l: border(side("border-left")),
        border_r: border(side("border-right")),
        border_t: border(side("border-top")),
        border_b: border(side("border-bottom")),
        has_direct_border_l: true,
        has_direct_border_r: true,
        has_direct_border_t: true,
        has_direct_border_b: true,
        has_direct_diagonal_tl: false,
        has_direct_diagonal_tr: false,
        diagonal_tl: border(cell_props.get(NS_STYLE, "diagonal-tl-br")),
        diagonal_tr: border(cell_props.get(NS_STYLE, "diagonal-bl-tr")),
        grid_span: 1,
        row_span: 1,
        h_merge: false,
        v_merge: false,
    })
}
