// Text-focused markdown projection for docx documents. Separate code path
// from the rich JSON serialization used by the viewer — lossy by design.
// Preserves headings (via outlineLevel), bullet / numbered lists, tables,
// footnote bodies, and rich-text formatting; discards geometry, section
// properties, font metrics, drawing shapes, page layout.

use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;

use crate::types::{
    BodyElement, CellElement, DocParagraph, DocRun, DocTable, DocTableCell, Document, TextRun,
};

pub(crate) fn render_document(doc: &Document) -> String {
    let mut out = String::new();
    render_body(&doc.body, &mut out);

    if !doc.footnotes.is_empty() {
        out.push_str("\n## Footnotes\n\n");
        for note in &doc.footnotes {
            let text = note_inline_text(&note.content);
            let text = text.trim();
            if text.is_empty() {
                continue;
            }
            let _ = writeln!(out, "[^{}]: {}", note.id, text);
        }
    }
    if !doc.endnotes.is_empty() {
        out.push_str("\n## Endnotes\n\n");
        for note in &doc.endnotes {
            let text = note_inline_text(&note.content);
            let text = text.trim();
            if text.is_empty() {
                continue;
            }
            let _ = writeln!(out, "[^en{}]: {}", note.id, text);
        }
    }
    if !doc.comments.is_empty() {
        render_review_comments(&doc.comments, &mut out);
    }
    out
}

fn render_review_comments(comments: &[crate::types::DocxComment], out: &mut String) {
    if comments.is_empty() {
        return;
    }
    out.push_str("\n## Review comments\n\n");

    let mut children = HashMap::<&str, Vec<usize>>::new();
    for (index, comment) in comments.iter().enumerate() {
        if let Some(parent_id) = comment.parent_id.as_deref() {
            children.entry(parent_id).or_default().push(index);
        }
    }

    let mut emitted = HashSet::new();
    for comment in comments
        .iter()
        .filter(|comment| comment.parent_id.is_none())
    {
        render_comment_thread(comment, comments, &children, &mut emitted, out);
    }
    // Malformed extension metadata may reference a missing parent or form a
    // cycle. Keep every comment visible once instead of guessing a repair.
    for comment in comments {
        if !emitted.contains(&comment.id) {
            render_comment_thread(comment, comments, &children, &mut emitted, out);
        }
    }
}

fn render_comment_thread(
    comment: &crate::types::DocxComment,
    comments: &[crate::types::DocxComment],
    children: &HashMap<&str, Vec<usize>>,
    emitted: &mut HashSet<String>,
    out: &mut String,
) {
    if !emitted.insert(comment.id.clone()) {
        return;
    }
    let status = if comment.resolved == Some(true) {
        " (resolved)"
    } else {
        ""
    };
    let _ = writeln!(
        out,
        "### Comment {}{}\n",
        escape_heading_label(&comment.id),
        status
    );
    write_quoted_comment(
        comment.author.as_deref().unwrap_or("(unknown)"),
        comment.text.trim(),
        ">",
        "",
        out,
    );

    let mut pending: Vec<(usize, usize)> = children
        .get(comment.id.as_str())
        .into_iter()
        .flatten()
        .rev()
        .map(|index| (*index, 2))
        .collect();
    while let Some((index, depth)) = pending.pop() {
        let reply = &comments[index];
        if !emitted.insert(reply.id.clone()) {
            continue;
        }
        let prefix = ">".repeat(depth);
        let status = if reply.resolved == Some(true) {
            " (resolved)"
        } else {
            ""
        };
        write_quoted_comment(
            reply.author.as_deref().unwrap_or("(unknown)"),
            reply.text.trim(),
            &prefix,
            status,
            out,
        );
        if let Some(grandchildren) = children.get(reply.id.as_str()) {
            pending.extend(
                grandchildren
                    .iter()
                    .rev()
                    .map(|child_index| (*child_index, depth + 1)),
            );
        }
    }
    out.push('\n');
}

fn write_quoted_comment(author: &str, text: &str, prefix: &str, status: &str, out: &mut String) {
    let author = escape_inline_md(&author.replace(['\r', '\n'], " "));
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

fn escape_heading_label(value: &str) -> String {
    value
        .replace(['\r', '\n'], " ")
        .replace('\\', "\\\\")
        .replace('#', "\\#")
}

fn render_body(body: &[BodyElement], out: &mut String) {
    for el in body {
        match el {
            BodyElement::Paragraph(p) => render_paragraph(p, out),
            BodyElement::Table(t) => render_table(t, out),
            BodyElement::PageBreak { .. }
            | BodyElement::ColumnBreak
            | BodyElement::SectionBreak { .. } => {
                // Page / column / section breaks are layout, not content — skip
                // in the projection.
            }
        }
    }
}

fn render_paragraph(p: &DocParagraph, out: &mut String) {
    let text = render_runs(&p.runs, &p.run_revisions);
    let trimmed = text.trim();
    if trimmed.is_empty() {
        out.push('\n');
        return;
    }

    // ECMA-376 §17.3.1.20 outlineLvl 0-8 → markdown `#`-`######`.
    // Levels 6-8 collapse to `######` (markdown caps at 6).
    if let Some(level) = p.outline_level {
        let hashes = "#".repeat(((level as usize) + 1).min(6));
        let _ = writeln!(out, "{} {}\n", hashes, trimmed);
        return;
    }

    // Numbering / bullets — `format` is the abstract num's level format
    // ("decimal" / "bullet" / "lowerLetter" / etc.). Bullet → `-`; everything
    // else (decimal / roman / letter) → `1.` and let the markdown renderer
    // auto-number sequential items.
    if let Some(num) = &p.numbering {
        let indent = "  ".repeat(num.level as usize);
        let marker = if num.format == "bullet" { "-" } else { "1." };
        let _ = writeln!(out, "{}{} {}", indent, marker, trimmed);
        return;
    }

    let _ = writeln!(out, "{}\n", trimmed);
}

/// Flatten a note's block-level content into a single inline markdown string
/// (paragraphs joined with a space). Used for the `[^id]:` footnote/endnote
/// projection, which is single-line by convention. Reference markers are
/// dropped by `format_text_run`.
fn note_inline_text(content: &[BodyElement]) -> String {
    let mut parts: Vec<String> = Vec::new();
    for el in content {
        if let BodyElement::Paragraph(p) = el {
            let t = render_runs(&p.runs, &p.run_revisions);
            let t = t.trim();
            if !t.is_empty() {
                parts.push(t.to_string());
            }
        }
    }
    parts.join(" ")
}

fn render_runs(runs: &[DocRun], revisions: &[Option<crate::types::RunRevision>]) -> String {
    let mut out = String::new();
    for (run_index, run) in runs.iter().enumerate() {
        // §17.13.5 revision containers apply to every inline occurrence. Keep
        // Markdown on the same accepted-final projection as layout, including
        // fields, math, breaks, tabs, and drawings rather than only text.
        let revision = revisions
            .get(run_index)
            .and_then(Option::as_ref)
            .or(match run {
                DocRun::Text(text) => text.revision.as_ref(),
                _ => None,
            });
        if revision.is_some_and(|value| value.kind == "deletion" || value.kind == "moveFrom") {
            continue;
        }
        match run {
            DocRun::Text(t) => out.push_str(&format_text_run(t)),
            DocRun::Field(f) => {
                // Field runs render their displayed text (PAGE, NUMPAGES, …
                // resolve at view time in the renderer; for markdown we just
                // surface whatever fallback the parser captured).
                if !f.fallback_text.is_empty() {
                    out.push_str(&escape_inline_md(&f.fallback_text));
                }
            }
            DocRun::Break { break_type } => {
                use crate::types::BreakType;
                match break_type {
                    BreakType::Line | BreakType::RenderedPage => out.push_str("  \n"),
                    BreakType::Page | BreakType::Column => out.push_str("\n\n"),
                }
            }
            DocRun::AnchorHost(_)
            | DocRun::Image(_)
            | DocRun::UnavailableDrawing(_)
            | DocRun::Shape(_)
            | DocRun::Chart(_) => {
                // No readable text; intentionally dropped. Use docx_get_images
                // / docx_get_shapes when you need metadata.
            }
            DocRun::Math { nodes, .. } => {
                // Surface the equation's literal characters as inline text.
                let text = crate::math::nodes_to_text(nodes);
                if !text.is_empty() {
                    out.push_str(&escape_inline_md(&text));
                }
            }
            DocRun::PTab { .. } => {
                // §17.3.3.23 absolute-position tab. Markdown has no notion of an
                // absolute tab position, so project it as a plain tab advance (the
                // same whitespace projection a `<w:tab>`'s "\t" text run produces).
                out.push('\t');
            }
        }
    }
    out
}

fn format_text_run(t: &crate::types::TextRun) -> String {
    // Footnote/endnote reference markers carry only the note's id; the linkage
    // is expressed via the `[^id]` syntax elsewhere, so drop the marker glyph
    // from the inline text projection.
    if t.note_ref.is_some() {
        return String::new();
    }
    let raw = &t.text;
    if raw.is_empty() {
        return String::new();
    }
    if raw.chars().all(|c| c.is_whitespace()) {
        return raw.clone();
    }
    // Pull whitespace outside the formatting wrappers so `(bold) " Title " `
    // becomes ` **Title** ` not `**" Title "**`.
    let leading_len = raw.len() - raw.trim_start().len();
    let trail_start = raw.trim_end().len();
    let leading = &raw[..leading_len];
    let trailing = &raw[trail_start..];
    let body = &raw[leading_len..trail_start];

    let mut s = escape_inline_md(body);
    if let Some(url) = &t.hyperlink {
        s = format!("[{s}]({url})");
    }
    // Order: bold > italic > strikethrough wrappers. Multiple wrappers stack.
    if t.bold {
        s = format!("**{s}**");
    }
    if t.italic {
        s = format!("*{s}*");
    }
    if t.strikethrough {
        s = format!("~~{s}~~");
    }
    let mut out = String::with_capacity(leading.len() + s.len() + trailing.len());
    out.push_str(leading);
    out.push_str(&s);
    out.push_str(trailing);
    out
}

/// Minimal markdown escape: only metacharacters that would otherwise be
/// parsed as formatting (bold `*`, italic `_`, code `` ` ``, backslash).
/// Pipes are handled separately inside table cells.
fn escape_inline_md(s: &str) -> String {
    s.replace('\\', "\\\\")
        .replace('*', "\\*")
        .replace('_', "\\_")
        .replace('`', "\\`")
}

fn render_table(t: &DocTable, out: &mut String) {
    if t.rows.is_empty() {
        return;
    }
    let cols = t.rows[0].cells.len();
    if cols == 0 {
        return;
    }
    let header_cells: Vec<String> = t.rows[0].cells.iter().map(render_table_cell).collect();
    let _ = writeln!(out, "| {} |", header_cells.join(" | "));
    let sep: Vec<&str> = (0..cols).map(|_| "---").collect();
    let _ = writeln!(out, "| {} |", sep.join(" | "));
    for row in t.rows.iter().skip(1) {
        let cells: Vec<String> = row.cells.iter().map(render_table_cell).collect();
        let _ = writeln!(out, "| {} |", cells.join(" | "));
    }
    out.push('\n');
}

fn render_table_cell(cell: &DocTableCell) -> String {
    // vMerge=continuation → leave empty so the row alignment stays intact.
    if matches!(cell.v_merge, Some(false)) {
        return String::new();
    }
    let mut buf = String::new();
    for (i, el) in cell.content.iter().enumerate() {
        if i > 0 {
            buf.push_str("<br>");
        }
        match el {
            CellElement::Paragraph(p) => {
                let text = render_runs(&p.runs, &p.run_revisions);
                buf.push_str(text.trim());
            }
            CellElement::Table(_) => {
                // Nested tables: not representable inside a markdown cell.
                // Skip — agents that need the structure should use
                // docx_get_table on the outer cell's paragraphs.
                buf.push_str("(nested table)");
            }
        }
    }
    buf.replace('|', "\\|")
}

// Silence unused-import warnings when the cfg gate excludes some types.
#[allow(dead_code)]
fn _ensure_types_used(_t: TextRun) {}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::DocxComment;

    #[test]
    fn review_comments_are_separated_quoted_and_threaded() {
        let comments = vec![
            DocxComment {
                id: "12".to_string(),
                author: Some("Reviewer".to_string()),
                initials: None,
                date: None,
                text: "Check this line.\nKeep the wording.".to_string(),
                parent_id: None,
                resolved: Some(true),
                paragraphs: Vec::new(),
            },
            DocxComment {
                id: "13".to_string(),
                author: Some("Editor".to_string()),
                initials: None,
                date: None,
                text: "Updated.".to_string(),
                parent_id: Some("12".to_string()),
                resolved: None,
                paragraphs: Vec::new(),
            },
        ];
        let mut out = String::new();

        render_review_comments(&comments, &mut out);

        assert!(out.starts_with("\n## Review comments\n"), "{out}");
        assert!(out.contains("### Comment 12 (resolved)"), "{out}");
        assert!(
            out.contains("> **Reviewer**\n>\n> Check this line.\n> Keep the wording."),
            "{out}"
        );
        assert!(out.contains(">> **Editor**\n>>\n>> Updated."), "{out}");
        assert_eq!(out.matches("### Comment").count(), 1, "{out}");
    }
}
