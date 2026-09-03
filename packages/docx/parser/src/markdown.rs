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
            let _ = writeln!(out, "[^{}]: {}", markdown_label(&note.id), text);
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
            let _ = writeln!(out, "[^en{}]: {}", markdown_label(&note.id), text);
        }
    }
    if !doc.comments.is_empty() {
        out.push_str("\n## Comments\n\n");
        let children = comment_children(&doc.comments);
        let mut emitted = HashSet::new();
        for comment in doc
            .comments
            .iter()
            .filter(|comment| comment.parent_id.is_none())
        {
            render_comment_definition(comment, &doc.comments, &children, &mut emitted, &mut out);
        }
        // Malformed extension metadata can point at a missing parent. Keep the
        // comment visible as its own definition rather than silently dropping it.
        for comment in &doc.comments {
            if !emitted.contains(&comment.id) {
                render_comment_definition(
                    comment,
                    &doc.comments,
                    &children,
                    &mut emitted,
                    &mut out,
                );
            }
        }
    }
    out
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
    let text = render_paragraph_runs(p);
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
/// projection, which is single-line by convention. The note body's own empty-id
/// auto-number placeholder is dropped by `format_text_run`.
fn note_inline_text(content: &[BodyElement]) -> String {
    let mut parts: Vec<String> = Vec::new();
    for el in content {
        if let BodyElement::Paragraph(p) = el {
            let t = render_paragraph_runs(p);
            let t = t.trim();
            if !t.is_empty() {
                parts.push(t.to_string());
            }
        }
    }
    parts.join(" ")
}

fn render_paragraph_runs(paragraph: &DocParagraph) -> String {
    if paragraph.comment_marks.is_empty() {
        return render_runs(&paragraph.runs, &paragraph.run_revisions);
    }
    let mut preferred = HashMap::<&str, &str>::new();
    for mark in &paragraph.comment_marks {
        let entry = preferred
            .entry(mark.id.as_str())
            .or_insert(mark.kind.as_str());
        if mark.kind == "reference" || (*entry != "reference" && mark.kind == "rangeEnd") {
            *entry = mark.kind.as_str();
        }
    }
    let mut at_boundary = HashMap::<usize, Vec<&str>>::new();
    for mark in &paragraph.comment_marks {
        if preferred.get(mark.id.as_str()).copied() == Some(mark.kind.as_str()) {
            let ids = at_boundary.entry(mark.run_index as usize).or_default();
            if !ids.contains(&mark.id.as_str()) {
                ids.push(mark.id.as_str());
            }
        }
    }

    let mut out = String::new();
    let mut emitted_ids = HashSet::new();
    for index in 0..=paragraph.runs.len() {
        if let Some(ids) = at_boundary.get(&index) {
            for id in ids {
                if emitted_ids.insert(*id) {
                    let _ = write!(out, "[^comment-{}]", markdown_label(id));
                }
            }
        }
        if index < paragraph.runs.len() {
            let revision = paragraph.run_revisions.get(index..=index).unwrap_or(&[]);
            out.push_str(&render_runs(&paragraph.runs[index..=index], revision));
        }
    }
    out
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
    if let Some(note) = &t.note_ref {
        // Empty ids are the auto-number placeholders inside note bodies, not
        // references from the document body.
        if note.id.is_empty() {
            return String::new();
        }
        let prefix = if note.kind == "endnote" { "en" } else { "" };
        return format!("[^{prefix}{}]", markdown_label(&note.id));
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
                let text = render_paragraph_runs(p);
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

fn markdown_label(value: &str) -> String {
    value
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() || matches!(character, '-' | '_') {
                character
            } else {
                '-'
            }
        })
        .collect()
}

fn comment_children(comments: &[crate::types::DocxComment]) -> HashMap<&str, Vec<usize>> {
    let mut children = HashMap::<&str, Vec<usize>>::new();
    for (index, comment) in comments.iter().enumerate() {
        if let Some(parent) = comment.parent_id.as_deref() {
            children.entry(parent).or_default().push(index);
        }
    }
    children
}

fn render_comment_definition(
    comment: &crate::types::DocxComment,
    comments: &[crate::types::DocxComment],
    children: &HashMap<&str, Vec<usize>>,
    emitted: &mut HashSet<String>,
    out: &mut String,
) {
    if !emitted.insert(comment.id.clone()) {
        return;
    }
    let author = comment.author.as_deref().unwrap_or("(unknown)");
    let status = if comment.resolved == Some(true) {
        " [resolved]"
    } else {
        ""
    };
    let body = comment.text.trim().replace('\n', "<br>");
    let _ = writeln!(
        out,
        "[^comment-{}]: **{}{}**: {}",
        markdown_label(&comment.id),
        author,
        status,
        body
    );
    render_comment_replies(&comment.id, comments, children, emitted, 1, out);
}

fn render_comment_replies(
    parent_id: &str,
    comments: &[crate::types::DocxComment],
    children: &HashMap<&str, Vec<usize>>,
    emitted: &mut HashSet<String>,
    depth: usize,
    out: &mut String,
) {
    let mut pending: Vec<(usize, usize)> = children
        .get(parent_id)
        .into_iter()
        .flatten()
        .rev()
        .map(|index| (*index, depth))
        .collect();
    while let Some((index, reply_depth)) = pending.pop() {
        let reply = &comments[index];
        if !emitted.insert(reply.id.clone()) {
            continue;
        }
        let author = reply.author.as_deref().unwrap_or("(unknown)");
        let status = if reply.resolved == Some(true) {
            " [resolved]"
        } else {
            ""
        };
        let indent = "    ".repeat(reply_depth);
        let body = reply.text.trim().replace('\n', "<br>");
        let _ = writeln!(out, "{indent}- **{author}{status}**: {body}");
        if let Some(grandchildren) = children.get(reply.id.as_str()) {
            pending.extend(
                grandchildren
                    .iter()
                    .rev()
                    .map(|index| (*index, reply_depth + 1)),
            );
        }
    }
}

// Silence unused-import warnings when the cfg gate excludes some types.
#[allow(dead_code)]
fn _ensure_types_used(_t: TextRun) {}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::{DocxCommentMark, NoteRef};

    #[test]
    fn note_references_link_to_the_definitions_emitted_at_document_end() {
        assert_eq!(
            format_text_run(&TextRun {
                note_ref: Some(NoteRef {
                    kind: "footnote".to_string(),
                    id: "7".to_string(),
                }),
                ..Default::default()
            }),
            "[^7]"
        );
        assert_eq!(
            format_text_run(&TextRun {
                note_ref: Some(NoteRef {
                    kind: "endnote".to_string(),
                    id: "2".to_string(),
                }),
                ..Default::default()
            }),
            "[^en2]"
        );
    }

    #[test]
    fn comment_reference_is_inserted_once_at_its_authored_boundary() {
        let paragraph = DocParagraph {
            runs: vec![DocRun::Text(Box::new(TextRun {
                text: "Reviewed text".to_string(),
                ..Default::default()
            }))],
            comment_marks: vec![
                DocxCommentMark {
                    id: "12".to_string(),
                    kind: "rangeEnd".to_string(),
                    run_index: 1,
                },
                DocxCommentMark {
                    id: "12".to_string(),
                    kind: "reference".to_string(),
                    run_index: 1,
                },
            ],
            ..Default::default()
        };

        assert_eq!(
            render_paragraph_runs(&paragraph),
            "Reviewed text[^comment-12]"
        );
    }
}
