// Text-focused markdown projection for xlsx workbooks. Walks a parsed
// Worksheet JSON value and emits a `## SheetName` heading followed by a
// one or more pipe tables containing the cells' cached display values. Designed for AI
// agents that need to read spreadsheet content efficiently — drops
// styling, formatting (numFmt), charts, sparklines, drawings, slicers,
// conditional formatting, and the formula text (cached value only).

use std::collections::HashSet;
use std::fmt::Write as _;

use serde_json::Value;

use crate::types::SharedString;

pub(crate) fn render_sheet(sheet: &Value, shared_strings: &[SharedString], out: &mut String) {
    let name = sheet["name"].as_str().unwrap_or("(unnamed)");
    let _ = writeln!(out, "## {}\n", escape_heading(name));

    let Some(rows) = sheet["rows"].as_array() else {
        render_comments(sheet, out);
        return;
    };

    // Find the populated bounding box. Rows are stored sparsely keyed by
    // 1-based row index; cells likewise carry their 1-based col. We render
    // the rectangle [min_row..=max_row] × [min_col..=max_col] dense, filling
    // gaps with empty cells.
    let mut min_row = u32::MAX;
    let mut max_row = 0u32;
    let mut min_col = u32::MAX;
    let mut max_col = 0u32;
    for row in rows {
        let row_idx = row["index"].as_u64().unwrap_or(0) as u32;
        let Some(cells) = row["cells"].as_array() else {
            continue;
        };
        for cell in cells {
            let s = cell_display(cell, shared_strings);
            if s.is_empty() {
                continue;
            }
            let col = cell["col"].as_u64().unwrap_or(0) as u32;
            if row_idx == 0 || col == 0 {
                continue;
            }
            if row_idx < min_row {
                min_row = row_idx;
            }
            if row_idx > max_row {
                max_row = row_idx;
            }
            if col < min_col {
                min_col = col;
            }
            if col > max_col {
                max_col = col;
            }
        }
    }
    if max_row == 0 || max_col == 0 {
        // Empty sheet — comments may still carry useful cell-attached review text.
        render_comments(sheet, out);
        return;
    }

    // Build a dense grid of display strings indexed by [r - min_row][c - min_col].
    let n_rows = (max_row - min_row + 1) as usize;
    let n_cols = (max_col - min_col + 1) as usize;
    let mut grid: Vec<Vec<String>> = vec![vec![String::new(); n_cols]; n_rows];
    for row in rows {
        let row_idx = row["index"].as_u64().unwrap_or(0) as u32;
        if row_idx < min_row || row_idx > max_row {
            continue;
        }
        let Some(cells) = row["cells"].as_array() else {
            continue;
        };
        for cell in cells {
            let col = cell["col"].as_u64().unwrap_or(0) as u32;
            if col < min_col || col > max_col {
                continue;
            }
            let r = (row_idx - min_row) as usize;
            let c = (col - min_col) as usize;
            grid[r][c] = cell_display(cell, shared_strings);
        }
    }

    // Apply merges: ECMA-376 §18.3.1.55 — the top-left cell carries the value,
    // continuation cells must render empty. We do this after grid population so
    // a value living at the top-left of a merge survives and the rest clear.
    let merge_continuation = collect_merge_continuation_cells(&sheet["mergeCells"]);
    for (row_idx, col) in merge_continuation {
        if row_idx < min_row || row_idx > max_row || col < min_col || col > max_col {
            continue;
        }
        grid[(row_idx - min_row) as usize][(col - min_col) as usize].clear();
    }

    let table_ranges = table_ranges(&sheet["tables"]);
    let regions = semantic_regions(&grid, min_row, min_col, &table_ranges);
    let multiple_regions = regions.len() > 1;
    for region in regions {
        if multiple_regions {
            let _ = writeln!(
                out,
                "### {}:{}\n",
                cell_ref(region.top_row, region.left_col),
                cell_ref(region.bottom_row, region.right_col)
            );
        }
        render_region(&grid, min_row, min_col, region, &table_ranges, out);
    }
    render_comments(sheet, out);
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Region {
    top_row: u32,
    bottom_row: u32,
    left_col: u32,
    right_col: u32,
}

#[derive(Clone, Copy, Debug)]
struct DeclaredTable {
    range: Region,
    has_header: bool,
}

fn semantic_regions(
    grid: &[Vec<String>],
    min_row: u32,
    min_col: u32,
    table_ranges: &[DeclaredTable],
) -> Vec<Region> {
    let initial = Region {
        top_row: min_row,
        bottom_row: min_row + grid.len() as u32 - 1,
        left_col: min_col,
        right_col: min_col + grid[0].len() as u32 - 1,
    };
    let mut out = Vec::new();
    let mut pending = vec![initial];
    while let Some(region) = pending.pop() {
        // Blank row bands are semantic section boundaries. A declared Excel
        // table remains atomic even when it contains a blank data row.
        let row_split = (region.top_row..=region.bottom_row).find_map(|start| {
            if !row_is_separator(grid, min_row, min_col, start, region, table_ranges) {
                return None;
            }
            let mut end = start;
            while end < region.bottom_row
                && row_is_separator(grid, min_row, min_col, end + 1, region, table_ranges)
            {
                end += 1;
            }
            Some((start, end))
        });
        if let Some((start, end)) = row_split {
            // Stack is LIFO: push the later block first for top-to-bottom output.
            if end < region.bottom_row {
                pending.push(Region {
                    top_row: end + 1,
                    ..region
                });
            }
            if start > region.top_row {
                pending.push(Region {
                    bottom_row: start - 1,
                    ..region
                });
            }
            continue;
        }

        // With no horizontal separator, split side-by-side blocks on fully
        // blank columns. Combined with the row-first pass this is row-major.
        let col_split = (region.left_col..=region.right_col).find_map(|start| {
            if !col_is_separator(grid, min_row, min_col, start, region, table_ranges) {
                return None;
            }
            let mut end = start;
            while end < region.right_col
                && col_is_separator(grid, min_row, min_col, end + 1, region, table_ranges)
            {
                end += 1;
            }
            Some((start, end))
        });
        if let Some((start, end)) = col_split {
            if end < region.right_col {
                pending.push(Region {
                    left_col: end + 1,
                    ..region
                });
            }
            if start > region.left_col {
                pending.push(Region {
                    right_col: start - 1,
                    ..region
                });
            }
            continue;
        }
        out.push(region);
    }
    out
}

fn render_region(
    grid: &[Vec<String>],
    min_row: u32,
    min_col: u32,
    region: Region,
    table_ranges: &[DeclaredTable],
    out: &mut String,
) {
    let n_cols = (region.right_col - region.left_col + 1) as usize;
    let rows: Vec<Vec<String>> = (region.top_row..=region.bottom_row)
        .map(|row| {
            (region.left_col..=region.right_col)
                .map(|col| grid[(row - min_row) as usize][(col - min_col) as usize].clone())
                .collect()
        })
        .collect();
    let declared_header = table_ranges.iter().any(|table| {
        table.has_header
            && table.range.top_row == region.top_row
            && table.range.left_col <= region.left_col
            && table.range.right_col >= region.right_col
    });
    // Outside a declared Excel table we do not guess that the first row is a
    // header: promoting ordinary data to `<th>` changes its meaning. Synthetic
    // column labels keep every authored row in the table body.
    let has_header = declared_header;
    if has_header {
        write_table_row(out, &rows[0], n_cols);
    } else {
        let headers: Vec<String> = (region.left_col..=region.right_col)
            .map(column_name)
            .collect();
        write_table_row(out, &headers, n_cols);
    }
    let sep: Vec<&str> = (0..n_cols).map(|_| "---").collect();
    let _ = writeln!(out, "| {} |", sep.join(" | "));
    for row in rows.iter().skip(usize::from(has_header)) {
        write_table_row(out, row, n_cols);
    }
    out.push('\n');
}

fn row_is_empty(
    grid: &[Vec<String>],
    min_row: u32,
    min_col: u32,
    row: u32,
    left: u32,
    right: u32,
) -> bool {
    (left..=right).all(|col| grid[(row - min_row) as usize][(col - min_col) as usize].is_empty())
}

fn row_is_separator(
    grid: &[Vec<String>],
    min_row: u32,
    min_col: u32,
    row: u32,
    region: Region,
    table_ranges: &[DeclaredTable],
) -> bool {
    row_is_empty(
        grid,
        min_row,
        min_col,
        row,
        region.left_col,
        region.right_col,
    ) && !table_ranges.iter().any(|table| {
        row >= table.range.top_row
            && row <= table.range.bottom_row
            && ranges_overlap(
                region.left_col,
                region.right_col,
                table.range.left_col,
                table.range.right_col,
            )
    })
}

fn col_is_empty(
    grid: &[Vec<String>],
    min_row: u32,
    min_col: u32,
    col: u32,
    top: u32,
    bottom: u32,
) -> bool {
    (top..=bottom).all(|row| grid[(row - min_row) as usize][(col - min_col) as usize].is_empty())
}

fn col_is_separator(
    grid: &[Vec<String>],
    min_row: u32,
    min_col: u32,
    col: u32,
    region: Region,
    table_ranges: &[DeclaredTable],
) -> bool {
    col_is_empty(
        grid,
        min_row,
        min_col,
        col,
        region.top_row,
        region.bottom_row,
    ) && !table_ranges.iter().any(|table| {
        col >= table.range.left_col
            && col <= table.range.right_col
            && ranges_overlap(
                region.top_row,
                region.bottom_row,
                table.range.top_row,
                table.range.bottom_row,
            )
    })
}

fn ranges_overlap(a_start: u32, a_end: u32, b_start: u32, b_end: u32) -> bool {
    a_start <= b_end && b_start <= a_end
}

fn table_ranges(tables: &Value) -> Vec<DeclaredTable> {
    tables
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|table| {
            let range = &table["range"];
            Some(DeclaredTable {
                range: Region {
                    top_row: range["top"].as_u64()? as u32,
                    bottom_row: range["bottom"].as_u64()? as u32,
                    left_col: range["left"].as_u64()? as u32,
                    right_col: range["right"].as_u64()? as u32,
                },
                has_header: table["headerRowCount"].as_u64().unwrap_or(1) > 0,
            })
        })
        .collect()
}

fn column_name(mut col: u32) -> String {
    let mut chars = Vec::new();
    while col > 0 {
        col -= 1;
        chars.push((b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    chars.iter().rev().collect()
}

fn cell_ref(row: u32, col: u32) -> String {
    format!("{}{}", column_name(col), row)
}

fn escape_heading(value: &str) -> String {
    value.replace('\\', "\\\\").replace('#', "\\#")
}

fn render_comments(sheet: &Value, out: &mut String) {
    let Some(comments) = sheet["comments"].as_array() else {
        return;
    };
    if comments.is_empty() {
        return;
    }
    out.push_str("### Cell comments\n\n");
    for comment in comments {
        let cell = comment["cellRef"].as_str().unwrap_or("(unknown cell)");
        let author = comment["author"].as_str().unwrap_or("(unknown)");
        let status = if comment["resolved"].as_bool() == Some(true) {
            " [resolved]"
        } else {
            ""
        };
        let _ = writeln!(out, "#### {} — {}{}\n", cell, author, status);
        let text = comment["rootText"]
            .as_str()
            .or_else(|| comment["text"].as_str())
            .unwrap_or("");
        write_quote(text, "> ", out);
        if let Some(replies) = comment["replies"].as_array() {
            for reply in replies {
                let author = reply["author"].as_str().unwrap_or("(unknown)");
                let status = if reply["resolved"].as_bool() == Some(true) {
                    " [resolved]"
                } else {
                    ""
                };
                let _ = writeln!(out, ">> **{}{}**", author, status);
                write_quote(reply["text"].as_str().unwrap_or(""), ">> ", out);
            }
        }
        out.push('\n');
    }
}

fn write_quote(value: &str, prefix: &str, out: &mut String) {
    if value.is_empty() {
        let _ = writeln!(out, "{prefix}");
    } else {
        for line in value.lines() {
            let _ = writeln!(out, "{prefix}{line}");
        }
    }
}

fn write_table_row(out: &mut String, row: &[String], n_cols: usize) {
    let cells: Vec<String> = (0..n_cols)
        .map(|i| row.get(i).map(|s| escape_cell(s)).unwrap_or_default())
        .collect();
    let _ = writeln!(out, "| {} |", cells.join(" | "));
}

fn escape_cell(s: &str) -> String {
    // Pipe is the only inline-table metachar in GFM cells. Newlines must also
    // be flattened or they break the row — collapse to a literal `<br>` so the
    // user sees the line structure.
    s.replace('|', "\\|").replace('\n', "<br>")
}

fn cell_display(cell: &Value, shared_strings: &[SharedString]) -> String {
    // CellValue has `rename_all = "camelCase"` so the JSON tag is lowercase
    // ("text"/"number"/...). PascalCase would silently never match — same
    // class of bug that hid pptx_extract_text earlier.
    let value = &cell["value"];
    match value["type"].as_str().unwrap_or("empty") {
        "text" => value["text"].as_str().unwrap_or("").to_string(),
        // A `t="s"` cell ships only an `si` index now; resolve it back to the
        // shared-string table's plain text (markdown drops runs anyway,
        // matching the `"text"` arm).
        "shared" => value["si"]
            .as_u64()
            .and_then(|i| shared_strings.get(i as usize))
            .map(|s| s.text.clone())
            .unwrap_or_default(),
        "number" => value["number"]
            .as_f64()
            .map(format_number)
            .unwrap_or_default(),
        "bool" => value["bool"]
            .as_bool()
            .map(|b| {
                if b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            })
            .unwrap_or_default(),
        "error" => value["error"].as_str().unwrap_or("#ERR").to_string(),
        _ => String::new(),
    }
}

fn format_number(n: f64) -> String {
    // Integer-valued doubles → integer form so 2025 doesn't show as 2025.0.
    if n.is_finite() && n.fract() == 0.0 && n.abs() < 1e15 {
        return format!("{}", n as i64);
    }
    // Round to 10 significant digits to mask IEEE-754 ULP noise (702.6
    // round-trips through XML as 702.5999999999999). Trim trailing zeros so
    // 702.6 doesn't render as 702.6000000000.
    let s = format!("{n:.10}");
    let trimmed = s.trim_end_matches('0').trim_end_matches('.').to_string();
    if trimmed.is_empty() {
        "0".to_string()
    } else {
        trimmed
    }
}

/// Returns the set of (row, col) coordinates that are continuation cells of a
/// merged range — i.e. every cell in `[top..=bottom] × [left..=right]` except
/// the top-left, which keeps its value.
fn collect_merge_continuation_cells(merge_cells: &Value) -> HashSet<(u32, u32)> {
    let mut set = HashSet::new();
    let Some(arr) = merge_cells.as_array() else {
        return set;
    };
    for m in arr {
        let top = m["top"].as_u64().unwrap_or(0) as u32;
        let left = m["left"].as_u64().unwrap_or(0) as u32;
        let bottom = m["bottom"].as_u64().unwrap_or(0) as u32;
        let right = m["right"].as_u64().unwrap_or(0) as u32;
        if top == 0 || left == 0 || bottom < top || right < left {
            continue;
        }
        for r in top..=bottom {
            for c in left..=right {
                if r == top && c == left {
                    continue;
                }
                set.insert((r, c));
            }
        }
    }
    set
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    /// A `{"type":"shared","si":N}` cell must resolve to the sharedStrings
    /// table's text — the wire ships only the index, so markdown has to look
    /// it up (mirrors the runtime `t="s"` path).
    #[test]
    fn shared_cell_resolves_against_table() {
        let shared = vec![
            SharedString {
                text: "Alpha".to_string(),
                runs: None,
                ..Default::default()
            },
            SharedString {
                text: "Beta".to_string(),
                runs: None,
                ..Default::default()
            },
        ];
        let sheet = json!({
            "name": "Sheet1",
            "rows": [
                {
                    "index": 1,
                    "cells": [
                        { "col": 1, "row": 1, "value": { "type": "shared", "si": 1 } },
                        { "col": 2, "row": 1, "value": { "type": "number", "number": 3.0 } }
                    ]
                }
            ]
        });
        let mut out = String::new();
        render_sheet(&sheet, &shared, &mut out);
        assert!(
            out.contains("Beta"),
            "si=1 must resolve to shared[1] text, got:\n{out}"
        );
    }

    /// An out-of-range `si` resolves to empty text (historical fallback),
    /// leaving no populated cell — the sheet renders as empty.
    #[test]
    fn shared_cell_out_of_range_is_empty() {
        let shared: Vec<SharedString> = Vec::new();
        let sheet = json!({
            "name": "Sheet1",
            "rows": [
                {
                    "index": 1,
                    "cells": [
                        { "col": 1, "row": 1, "value": { "type": "shared", "si": 9 } }
                    ]
                }
            ]
        });
        let mut out = String::new();
        render_sheet(&sheet, &shared, &mut out);
        // No populated cells → only the heading is emitted (no table body).
        assert!(!out.contains('|'), "empty sheet must have no table: {out}");
    }

    #[test]
    fn blank_rows_and_columns_create_separate_row_major_regions() {
        let sheet = json!({
            "name": "Blocks",
            "rows": [
                { "index": 1, "cells": [
                    { "col": 1, "value": { "type": "text", "text": "top-left" } },
                    { "col": 3, "value": { "type": "text", "text": "top-right" } }
                ]},
                { "index": 3, "cells": [
                    { "col": 1, "value": { "type": "text", "text": "bottom-left" } },
                    { "col": 3, "value": { "type": "text", "text": "bottom-right" } }
                ]}
            ]
        });
        let mut out = String::new();
        render_sheet(&sheet, &[], &mut out);

        let positions: Vec<usize> = ["top-left", "top-right", "bottom-left", "bottom-right"]
            .iter()
            .map(|value| out.find(value).expect("region value"))
            .collect();
        assert!(positions.windows(2).all(|pair| pair[0] < pair[1]), "{out}");
        assert!(out.contains("### A1:A1"), "{out}");
        assert!(out.contains("### C3:C3"), "{out}");
    }

    #[test]
    fn declared_table_keeps_blank_rows_and_uses_its_real_header() {
        let sheet = json!({
            "name": "Table",
            "rows": [
                { "index": 1, "cells": [
                    { "col": 1, "value": { "type": "text", "text": "Name" } }
                ]},
                { "index": 3, "cells": [
                    { "col": 1, "value": { "type": "text", "text": "Ada" } }
                ]}
            ],
            "tables": [{
                "range": { "top": 1, "left": 1, "bottom": 3, "right": 1 },
                "headerRowCount": 1
            }]
        });
        let mut out = String::new();
        render_sheet(&sheet, &[], &mut out);

        assert!(
            !out.contains("### A1:A1"),
            "declared table must stay atomic: {out}"
        );
        assert!(out.contains("| Name |\n| --- |\n|  |\n| Ada |"), "{out}");
    }

    #[test]
    fn comments_are_collected_after_the_sheet_data_with_thread_context() {
        let sheet = json!({
            "name": "Review",
            "rows": [{ "index": 1, "cells": [
                { "col": 1, "value": { "type": "text", "text": "Value" } }
            ]}],
            "comments": [{
                "cellRef": "A1",
                "author": "Mina",
                "rootText": "Check this",
                "text": "Check this Reply",
                "replies": [{ "author": "Ren", "text": "Done" }]
            }]
        });
        let mut out = String::new();
        render_sheet(&sheet, &[], &mut out);

        assert!(
            out.find("Value").unwrap() < out.find("### Cell comments").unwrap(),
            "{out}"
        );
        assert!(out.contains("#### A1 — Mina"), "{out}");
        assert!(out.contains(">> **Ren**\n>> Done"), "{out}");
    }
}
