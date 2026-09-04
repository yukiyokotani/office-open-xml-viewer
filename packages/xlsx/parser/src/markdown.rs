// Text-focused markdown projection for xlsx workbooks. Walks a parsed
// Worksheet JSON value and emits a `## SheetName` heading followed by a
// pipe table containing the cells' cached display values. Designed for AI
// agents that need to read spreadsheet content efficiently — drops
// styling, formatting (numFmt), charts, sparklines, drawings, slicers,
// conditional formatting, and the formula text (cached value only).

use std::collections::HashSet;
use std::fmt::Write as _;

use serde_json::Value;

use crate::types::SharedString;

pub(crate) fn render_sheet(sheet: &Value, shared_strings: &[SharedString], out: &mut String) {
    let name = sheet["name"].as_str().unwrap_or("(unnamed)");
    let _ = writeln!(out, "## {}\n", name);

    let Some(rows) = sheet["rows"].as_array() else {
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
        // Empty sheet — emit the heading only.
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

    // Header row: use the first row of the bbox. Markdown tables require a
    // header — if the first row is blank we still emit empty headers so
    // downstream renderers parse the rest correctly.
    write_table_row(out, &grid[0], n_cols);
    let sep: Vec<&str> = (0..n_cols).map(|_| "---").collect();
    let _ = writeln!(out, "| {} |", sep.join(" | "));
    for row in grid.iter().skip(1) {
        // Skip fully-empty middle rows to keep output tight.
        if row.iter().all(|c| c.is_empty()) {
            continue;
        }
        write_table_row(out, row, n_cols);
    }
    out.push('\n');
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

pub(crate) fn render_sheet_comments(sheet: &Value, out: &mut String, has_comments: &mut bool) {
    let Some(comments) = sheet["comments"].as_array() else {
        return;
    };
    if comments.is_empty() {
        return;
    }
    if !*has_comments {
        out.push_str("\n## Review comments\n\n");
        *has_comments = true;
    }
    let sheet_name = sheet["name"].as_str().unwrap_or("(unnamed)");
    for comment in comments {
        let cell = comment["cellRef"].as_str().unwrap_or("(unknown cell)");
        let status = if comment["resolved"].as_bool() == Some(true) {
            " (resolved)"
        } else {
            ""
        };
        let _ = writeln!(
            out,
            "### {} — {}{}\n",
            escape_heading_text(sheet_name),
            escape_heading_text(cell),
            status
        );
        let text = comment["rootText"]
            .as_str()
            .or_else(|| comment["text"].as_str())
            .unwrap_or("");
        write_quoted_comment(
            comment["author"].as_str().unwrap_or("(unknown)"),
            text,
            ">",
            "",
            out,
        );
        if let Some(replies) = comment["replies"].as_array() {
            for reply in replies {
                let status = if reply["resolved"].as_bool() == Some(true) {
                    " (resolved)"
                } else {
                    ""
                };
                write_quoted_comment(
                    reply["author"].as_str().unwrap_or("(unknown)"),
                    reply["text"].as_str().unwrap_or(""),
                    ">>",
                    status,
                    out,
                );
            }
        }
        out.push('\n');
    }
}

fn write_quoted_comment(author: &str, text: &str, prefix: &str, status: &str, out: &mut String) {
    let author = escape_inline_text(&author.replace(['\r', '\n'], " "));
    let _ = writeln!(out, "{prefix} **{author}**{status}");
    let _ = writeln!(out, "{prefix}");
    if text.is_empty() {
        let _ = writeln!(out, "{prefix}");
    } else {
        for line in text.lines() {
            let _ = writeln!(out, "{prefix} {line}");
        }
    }
}

fn escape_heading_text(value: &str) -> String {
    value
        .replace(['\r', '\n'], " ")
        .replace('\\', "\\\\")
        .replace('#', "\\#")
}

fn escape_inline_text(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('*', "\\*")
        .replace('_', "\\_")
        .replace('`', "\\`")
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
    fn review_comments_are_separated_quoted_and_threaded() {
        let sheet = json!({
            "name": "Review",
            "comments": [{
                "cellRef": "B4",
                "author": "Reviewer",
                "rootText": "Check this cell.\nKeep the formula.",
                "text": "Check this cell.\nKeep the formula.\nUpdated.",
                "resolved": true,
                "replies": [{ "author": "Editor", "text": "Updated." }]
            }]
        });
        let mut out = String::new();
        let mut has_comments = false;

        render_sheet_comments(&sheet, &mut out, &mut has_comments);

        assert!(out.starts_with("\n## Review comments\n"), "{out}");
        assert!(out.contains("### Review — B4 (resolved)"), "{out}");
        assert!(
            out.contains("> **Reviewer**\n>\n> Check this cell.\n> Keep the formula."),
            "{out}"
        );
        assert!(out.contains(">> **Editor**\n>>\n>> Updated."), "{out}");
    }
}
