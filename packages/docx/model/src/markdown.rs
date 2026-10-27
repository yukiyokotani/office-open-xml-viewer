// Text-focused markdown projection for docx documents. Separate code path
// from the rich JSON serialization used by the viewer — lossy by design.
// Preserves headings (via outlineLevel), bullet / numbered lists, tables,
// footnote bodies, and rich-text formatting; discards geometry, section
// properties, font metrics, drawing shapes, page layout.
//
// Lives in the model crate so every producer of this model — the OOXML DOCX
// parser and the legacy binary DOC direct reader alike — projects the same
// markdown from the same document facts.

use std::collections::{HashMap, HashSet};
use std::fmt::{self, Write as _};

use ooxml_common::markdown::{
    escape_fragment, escape_text, line_start_after_markup, split_lines, TextContext,
};

use crate::{
    BodyElement, CellElement, DocParagraph, DocRun, DocTable, DocTableCell, Document, TextRun,
};

/// UTF-8 markdown sink that stops retaining bytes at the configured ceiling.
/// Public conversion paths report the crossing through the package resource
/// governor; this writer prevents the oversized projection from being
/// materialized before that report can be produced.
pub struct MarkdownWriter {
    value: String,
    limit: u64,
    observed: u64,
    exceeded: bool,
}

impl MarkdownWriter {
    pub fn new(limit: u64) -> Self {
        Self {
            value: String::new(),
            limit,
            observed: 0,
            exceeded: false,
        }
    }

    pub fn observed(&self) -> u64 {
        self.observed
    }

    /// Whether the ceiling was crossed: retained output stops at the crossing
    /// while `observed` keeps counting the full projection size.
    pub fn exceeded(&self) -> bool {
        self.exceeded
    }

    pub fn into_string(self) -> String {
        self.value
    }

    fn push_str(&mut self, value: &str) {
        if self.exceeded {
            return;
        }
        self.observed = self
            .observed
            .saturating_add(u64::try_from(value.len()).unwrap_or(u64::MAX));
        if self.observed > self.limit {
            self.exceeded = true;
            return;
        }
        self.value.push_str(value);
    }

    fn push(&mut self, value: char) {
        let mut encoded = [0; 4];
        self.push_str(value.encode_utf8(&mut encoded));
    }
}

impl fmt::Write for MarkdownWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.push_str(value);
        if self.exceeded {
            Err(fmt::Error)
        } else {
            Ok(())
        }
    }
}

/// Unbounded projection for tests and internal callers that already hold a
/// bounded model. Production conversion paths must use
/// [`render_document_with_limit`] and report the crossing through their
/// resource governor.
pub fn render_document(doc: &Document) -> String {
    render_document_with_limit(doc, u64::MAX).into_string()
}

pub fn render_document_with_limit(doc: &Document, limit: u64) -> MarkdownWriter {
    let mut out = MarkdownWriter::new(limit);
    render_body(&doc.body, &mut out);

    if !doc.footnotes.is_empty() {
        out.push_str("\n## Footnotes\n\n");
        for note in &doc.footnotes {
            let text = note_inline_text(&note.content);
            if text.is_empty() {
                continue;
            }
            let _ = writeln!(out, "[^{}]: {}", escape_inline_md(&note.id), text);
        }
    }
    if !doc.endnotes.is_empty() {
        out.push_str("\n## Endnotes\n\n");
        for note in &doc.endnotes {
            let text = note_inline_text(&note.content);
            if text.is_empty() {
                continue;
            }
            let _ = writeln!(out, "[^en{}]: {}", escape_inline_md(&note.id), text);
        }
    }
    if !doc.comments.is_empty() {
        render_review_comments(&doc.comments, &mut out);
    }
    out
}

fn render_review_comments(comments: &[crate::DocxComment], out: &mut MarkdownWriter) {
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
        if out.exceeded() {
            return;
        }
    }
    // Malformed extension metadata may reference a missing parent or form a
    // cycle. Keep every comment visible once instead of guessing a repair.
    for comment in comments {
        if !emitted.contains(&comment.id) {
            render_comment_thread(comment, comments, &children, &mut emitted, out);
            if out.exceeded() {
                return;
            }
        }
    }
}

fn render_comment_thread(
    comment: &crate::DocxComment,
    comments: &[crate::DocxComment],
    children: &HashMap<&str, Vec<usize>>,
    emitted: &mut HashSet<String>,
    out: &mut MarkdownWriter,
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
        escape_text(&comment.id, TextContext::Heading),
        status
    );
    write_quoted_comment(
        comment.author.as_deref().unwrap_or("(unknown)"),
        comment.text.trim(),
        1,
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
        let status = if reply.resolved == Some(true) {
            " (resolved)"
        } else {
            ""
        };
        write_quoted_comment(
            reply.author.as_deref().unwrap_or("(unknown)"),
            reply.text.trim(),
            depth,
            status,
            out,
        );
        if out.exceeded() {
            break;
        }
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

fn write_quote_prefix(depth: usize, out: &mut MarkdownWriter) -> fmt::Result {
    for _ in 0..depth {
        out.write_char('>')?;
    }
    Ok(())
}

fn write_quoted_comment(
    author: &str,
    text: &str,
    depth: usize,
    status: &str,
    out: &mut MarkdownWriter,
) {
    let author = escape_inline_md(&author.replace(['\r', '\n'], " "));
    let _ = write_quote_prefix(depth, out).and_then(|()| writeln!(out, " **{author}**{status}"));
    let _ = write_quote_prefix(depth, out).and_then(|()| writeln!(out));
    if text.is_empty() {
        let _ = write_quote_prefix(depth, out).and_then(|()| writeln!(out));
    } else {
        for line in split_lines(text) {
            if write_quote_prefix(depth, out)
                .and_then(|()| writeln!(out, " {}", escape_inline_md(line)))
                .is_err()
            {
                break;
            }
        }
    }
}

fn render_body(body: &[BodyElement], out: &mut MarkdownWriter) {
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
        if out.exceeded() {
            return;
        }
    }
}

fn render_paragraph(p: &DocParagraph, out: &mut MarkdownWriter) {
    let context = if p.outline_level.is_some() {
        TextContext::Heading
    } else {
        TextContext::Inline
    };
    let trimmed = render_runs(&p.runs, &p.run_revisions, context);
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
/// dropped by `push_text_run`.
fn note_inline_text(content: &[BodyElement]) -> String {
    let mut parts: Vec<String> = Vec::new();
    for el in content {
        if let BodyElement::Paragraph(p) = el {
            let t = render_runs(&p.runs, &p.run_revisions, TextContext::Inline);
            if !t.is_empty() {
                parts.push(t);
            }
        }
    }
    parts.join(" ")
}

/// One piece of a projected paragraph: document text still to be escaped, or
/// Markdown the converter emits itself (wrappers, link syntax, breaks).
enum Piece {
    Text(String),
    Markup(String),
}

impl Piece {
    fn is_blank(&self) -> bool {
        match self {
            Piece::Text(value) | Piece::Markup(value) => value.trim().is_empty(),
        }
    }
}

/// Project a paragraph's runs for the Markdown construct `context` places
/// them in: document text goes through the shared escaping, and line breaks
/// follow the construct (a space in a heading, `<br>` in a table cell).
/// Surrounding whitespace is trimmed from the document text before it is
/// escaped, never from the escaped output, so a trim cannot expose a
/// line-start marker the escaping did not see.
fn render_runs(
    runs: &[DocRun],
    revisions: &[Option<crate::RunRevision>],
    context: TextContext,
) -> String {
    let mut pieces = Vec::new();
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
            DocRun::Text(t) => push_text_run(t, context, &mut pieces),
            DocRun::Field(f) => {
                // Field runs render their displayed text (PAGE, NUMPAGES, …
                // resolve at view time in the renderer; for markdown we just
                // surface whatever fallback the parser captured).
                if !f.fallback_text.is_empty() {
                    pieces.push(Piece::Text(f.fallback_text.clone()));
                }
            }
            DocRun::Break { break_type } => {
                use crate::BreakType;
                pieces.push(Piece::Markup(
                    match (context, break_type) {
                        (TextContext::Heading, _) => " ",
                        (TextContext::TableCell, _) => "<br>",
                        (_, BreakType::Line | BreakType::RenderedPage) => "  \n",
                        (_, BreakType::Page | BreakType::Column) => "\n\n",
                    }
                    .to_string(),
                ));
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
                let text = ooxml_common::math::nodes_to_text(nodes);
                if !text.is_empty() {
                    pieces.push(Piece::Text(text));
                }
            }
            DocRun::PTab { .. } => {
                // §17.3.3.23 absolute-position tab. Markdown has no notion of an
                // absolute tab position, so project it as a plain tab advance (the
                // same whitespace projection a `<w:tab>`'s "\t" text run produces).
                pieces.push(Piece::Text("\t".to_string()));
            }
        }
    }

    // Trim the paragraph's surrounding whitespace on the unescaped pieces.
    let start = pieces.iter().position(|piece| !piece.is_blank());
    let end = pieces.iter().rposition(|piece| !piece.is_blank());
    let (Some(start), Some(end)) = (start, end) else {
        return String::new();
    };
    let pieces = &mut pieces[start..=end];
    if let Some(Piece::Text(first)) = pieces.first_mut() {
        *first = first.trim_start().to_string();
    }
    if let Some(Piece::Text(last)) = pieces.last_mut() {
        *last = last.trim_end().to_string();
    }
    // Escape the pieces as one line sequence, so indentation that opens a
    // projected line is dropped even when it arrives in its own run.
    let mut out = String::new();
    let mut line_start = true;
    for piece in pieces.iter() {
        match piece {
            Piece::Text(text) => out.push_str(&escape_fragment(text, context, &mut line_start)),
            Piece::Markup(markup) => {
                out.push_str(markup);
                line_start = line_start_after_markup(markup, line_start);
            }
        }
    }
    out
}

fn push_text_run(t: &TextRun, context: TextContext, pieces: &mut Vec<Piece>) {
    // Footnote/endnote reference markers carry only the note's id; the linkage
    // is expressed via the `[^id]` syntax elsewhere, so drop the marker glyph
    // from the inline text projection.
    if t.note_ref.is_some() {
        return;
    }
    let raw = &t.text;
    if raw.is_empty() {
        return;
    }
    if raw.chars().all(|c| c.is_whitespace()) {
        pieces.push(Piece::Text(raw.clone()));
        return;
    }
    // Pull whitespace outside the formatting wrappers so `(bold) " Title " `
    // becomes ` **Title** ` not `**" Title "**`.
    let leading_len = raw.len() - raw.trim_start().len();
    let trail_start = raw.trim_end().len();
    let leading = &raw[..leading_len];
    let trailing = &raw[trail_start..];
    let body = &raw[leading_len..trail_start];

    // A target outside the viewer's hyperlink scheme policy (`javascript:`,
    // `data:`, ...) keeps its text but does not become a Markdown link.
    let link = t.hyperlink.as_deref().and_then(|url| {
        let url = ooxml_common::markdown::link_destination(url)?;
        // A `|` would end a table cell even inside a link destination.
        Some(if context == TextContext::TableCell {
            url.replace('|', "%7C")
        } else {
            url
        })
    });
    // Wrappers nest as strikethrough > italic > bold > link around the text.
    let mut open = String::new();
    let mut close = String::new();
    if t.strikethrough {
        open.push_str("~~");
    }
    if t.italic {
        open.push('*');
    }
    if t.bold {
        open.push_str("**");
    }
    if link.is_some() {
        open.push('[');
    }
    if let Some(url) = &link {
        close.push_str("](");
        close.push_str(url);
        close.push(')');
    }
    if t.bold {
        close.push_str("**");
    }
    if t.italic {
        close.push('*');
    }
    if t.strikethrough {
        close.push_str("~~");
    }
    if !leading.is_empty() {
        pieces.push(Piece::Text(leading.to_string()));
    }
    if !open.is_empty() {
        pieces.push(Piece::Markup(open));
    }
    pieces.push(Piece::Text(body.to_string()));
    if !close.is_empty() {
        pieces.push(Piece::Markup(close));
    }
    if !trailing.is_empty() {
        pieces.push(Piece::Text(trailing.to_string()));
    }
}

/// Document text in running Markdown; see
/// [`ooxml_common::markdown::escape_text`].
fn escape_inline_md(s: &str) -> String {
    escape_text(s, TextContext::Inline)
}

fn render_table(t: &DocTable, out: &mut MarkdownWriter) {
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
                let text = render_runs(&p.runs, &p.run_revisions, TextContext::TableCell);
                buf.push_str(&text);
            }
            CellElement::Table(_) => {
                // Nested tables: not representable inside a markdown cell.
                // Skip — agents that need the structure should use
                // docx_get_table on the outer cell's paragraphs.
                buf.push_str("(nested table)");
            }
        }
    }
    buf
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::DocxComment;

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
        let mut out = MarkdownWriter::new(u64::MAX);

        render_review_comments(&comments, &mut out);
        let out = out.into_string();

        assert!(out.starts_with("\n## Review comments\n"), "{out}");
        assert!(out.contains("### Comment 12 (resolved)"), "{out}");
        assert!(
            out.contains("> **Reviewer**\n>\n> Check this line.\n> Keep the wording."),
            "{out}"
        );
        assert!(out.contains(">> **Editor**\n>>\n>> Updated."), "{out}");
        assert_eq!(out.matches("### Comment").count(), 1, "{out}");
    }

    #[test]
    fn deep_review_thread_stops_at_the_shared_markdown_ceiling() {
        let comments = (0..20_000)
            .map(|index| DocxComment {
                id: index.to_string(),
                author: Some("Reviewer".to_string()),
                initials: None,
                date: None,
                text: String::new(),
                parent_id: (index > 0).then(|| (index - 1).to_string()),
                resolved: None,
                paragraphs: Vec::new(),
            })
            .collect::<Vec<_>>();
        let limit = 4_096;

        let output = render_document_with_limit(
            &Document {
                comments,
                ..Document::default()
            },
            limit,
        );

        assert!(output.observed() > limit);
        assert!(output.value.len() as u64 <= limit);
    }

    #[test]
    fn writer_counts_utf8_bytes_and_stops_at_the_first_crossing() {
        let mut output = MarkdownWriter::new(3);
        output.push_str("é");
        output.push_str("é");
        output.push_str("ignored after crossing");

        assert_eq!(output.observed(), 4);
        assert_eq!(output.into_string(), "é");
    }
}
