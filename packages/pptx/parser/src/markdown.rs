//! The PPTX → GitHub-flavoured-markdown projection. All public conversion paths
//! feed canonical slides through this one bounded writer and renderer.

use crate::markdown_layout::{project_slide, shape_typography};
use crate::types::*;
use ooxml_common::math::nodes_to_text;
use std::cell::Cell;
use std::fmt;
use std::rc::Rc;

/// A local heading must be followed by content within four of its own box
/// heights. This scale-relative bound works across slide sizes and avoids
/// linking a prominent label to an unrelated object in a distant region.
const HEADING_MAX_VERTICAL_GAP_IN_OWN_HEIGHTS: i64 = 4;
const HEADING_MAX_CHARACTERS: usize = 140;
const HEADING_FONT_SIZE_RATIO: f64 = 1.2;

/// UTF-8 markdown sink that stops retaining bytes at the configured ceiling.
/// The parser reports the crossing through its package-scoped limit reporter;
/// this type's single responsibility is preventing an oversized projection from
/// first becoming an oversized allocation.
pub(crate) struct MarkdownWriter {
    value: String,
    budget: Rc<MarkdownBudget>,
}

struct MarkdownBudget {
    limit: u64,
    observed: Cell<u64>,
    exceeded: Cell<bool>,
}

impl MarkdownWriter {
    pub(crate) fn new(limit: u64) -> Self {
        Self {
            value: String::new(),
            budget: Rc::new(MarkdownBudget {
                limit,
                observed: Cell::new(0),
                exceeded: Cell::new(false),
            }),
        }
    }

    /// Create another sink charged to the same total allocation budget. The
    /// comment appendix can then be assembled out of order without allowing
    /// narrative + appendix memory to grow to twice the configured limit.
    pub(crate) fn sharing_budget(other: &Self) -> Self {
        Self {
            value: String::new(),
            budget: Rc::clone(&other.budget),
        }
    }

    pub(crate) fn observed(&self) -> u64 {
        self.budget.observed.get()
    }

    pub(crate) fn into_string(self) -> String {
        self.value
    }

    pub(crate) fn push_str(&mut self, value: &str) {
        if self.budget.exceeded.get() {
            return;
        }
        let observed = self
            .budget
            .observed
            .get()
            .saturating_add(u64::try_from(value.len()).unwrap_or(u64::MAX));
        self.budget.observed.set(observed);
        if observed > self.budget.limit {
            self.budget.exceeded.set(true);
            return;
        }
        self.value.push_str(value);
    }

    /// Join text already charged through a writer sharing this budget.
    pub(crate) fn append_precounted(&mut self, value: String) {
        if !self.budget.exceeded.get() {
            self.value.push_str(&value);
        }
    }

    pub(crate) fn push(&mut self, value: char) {
        let mut encoded = [0; 4];
        self.push_str(value.encode_utf8(&mut encoded));
    }
}

impl fmt::Write for MarkdownWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.push_str(value);
        if self.budget.exceeded.get() {
            Err(fmt::Error)
        } else {
            Ok(())
        }
    }
}

pub(crate) fn render_slide_md(
    slide: &Slide,
    slide_width: i64,
    slide_height: i64,
    out: &mut MarkdownWriter,
) {
    use std::fmt::Write as _;
    let semantic = project_slide(slide, slide_width, slide_height);
    if let Some((_, title)) = &semantic.title {
        out.push_str("# ");
        write_escaped_inline_md(title, out);
        let _ = writeln!(out, " (slide {})\n", slide.slide_number);
    } else {
        let _ = writeln!(out, "# Slide {}\n", slide.slide_number);
    }
    for (block_index, block) in semantic.blocks.iter().enumerate() {
        if block.starts_new_region && block_index > 0 {
            out.push_str("---\n\n");
        }
        let heading_index = if block.related {
            inferred_block_heading(slide, &block.element_indices)
        } else if block.element_indices.len() == 1 {
            semantic
                .blocks
                .get(block_index + 1)
                .and_then(|next| next.element_indices.first())
                .copied()
                .filter(|next| inferred_standalone_heading(slide, block.element_indices[0], *next))
                .map(|_| block.element_indices[0])
        } else {
            None
        };
        for index in &block.element_indices {
            if Some(*index) == heading_index {
                render_shape_heading_md(&slide.elements[*index], out);
            } else {
                render_element_md(&slide.elements[*index], out);
            }
        }
    }
    if let Some(notes) = &slide.notes {
        let trimmed = notes.trim();
        if !trimmed.is_empty() {
            let _ = writeln!(out, "## Speaker notes\n\n{}\n", trimmed);
        }
    }
}

fn inferred_standalone_heading(slide: &Slide, index: usize, next_index: usize) -> bool {
    let SlideElement::Shape(shape) = &slide.elements[index] else {
        return false;
    };
    let Some(body) = &shape.text_body else {
        return false;
    };
    if body.paragraphs.len() != 1 {
        return false;
    }
    let paragraph = &body.paragraphs[0];
    let inherit_means_bullet = matches!(
        shape.placeholder_type.as_deref(),
        Some("body" | "subTitle" | "obj" | "tx")
    );
    if !matches!(
        paragraph_kind(&paragraph.bullet, inherit_means_bullet),
        ParaKind::Plain
    ) {
        return false;
    }
    let Some(text) = shape_text_plain(shape) else {
        return false;
    };
    if text.chars().count() > HEADING_MAX_CHARACTERS || !text.chars().any(char::is_alphabetic) {
        return false;
    }
    let typography = shape_typography(shape);
    let bold = shape_is_bold(shape);
    let mut sizes: Vec<f64> = slide
        .elements
        .iter()
        .filter_map(|element| match element {
            SlideElement::Shape(shape)
                if !matches!(
                    shape.placeholder_type.as_deref(),
                    Some("title" | "ctrTitle" | "sldNum" | "dt" | "ftr" | "hdr")
                ) =>
            {
                shape_typography(shape).map(|value| value.0)
            }
            _ => None,
        })
        .collect();
    let size_is_salient = typography.is_some_and(|(size, _)| {
        if sizes.len() < 2 {
            return false;
        }
        sizes.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
        let median = sizes[(sizes.len() - 1) / 2].max(1.0);
        size >= median * HEADING_FONT_SIZE_RATIO || (bold && size >= median)
    });
    let bold_is_salient = bold
        && match &slide.elements[next_index] {
            SlideElement::Shape(next) => !shape_is_bold(next),
            _ => true,
        };
    if !size_is_salient && !bold_is_salient {
        return false;
    }

    let (x, y, width, height) = element_bounds(&slide.elements[index]);
    let (next_x, next_y, next_width, _) = element_bounds(&slide.elements[next_index]);
    if next_y < y || width <= 0 || next_width <= 0 {
        return false;
    }
    let overlap = x
        .saturating_add(width)
        .min(next_x.saturating_add(next_width))
        .saturating_sub(x.max(next_x));
    let same_column = overlap.saturating_mul(2) >= width.min(next_width);
    // A heading should introduce nearby content, not merely precede an
    // unrelated object elsewhere in the same broad column.
    let nearby = next_y.saturating_sub(y.saturating_add(height))
        <= height
            .max(1)
            .saturating_mul(HEADING_MAX_VERTICAL_GAP_IN_OWN_HEIGHTS);
    same_column && nearby
}

pub(crate) fn shape_text_plain(s: &ShapeElement) -> Option<String> {
    let tb = s.text_body.as_ref()?;
    let mut buf = String::new();
    for para in &tb.paragraphs {
        for run in &para.runs {
            if let TextRun::Text(t) = run {
                buf.push_str(&t.text);
            }
        }
        buf.push(' ');
    }
    let trimmed = buf.trim().to_string();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed)
    }
}

fn inferred_block_heading(slide: &Slide, indices: &[usize]) -> Option<usize> {
    if indices.len() < 2 {
        return None;
    }
    // A compact label deliberately overlapping the top edge of a larger
    // content-bearing panel is a common visual heading treatment even when
    // both shapes use the same font size and weight.
    if let Some(index) = indices
        .iter()
        .copied()
        .find(|index| attached_panel_label(slide, *index, indices))
    {
        return Some(index);
    }
    let mut styled: Vec<(usize, f64, bool)> = indices
        .iter()
        .filter_map(|index| {
            let SlideElement::Shape(shape) = &slide.elements[*index] else {
                return None;
            };
            let text = shape_text_plain(shape)?;
            let single_paragraph = shape.text_body.as_ref()?.paragraphs.len() == 1;
            let not_numeric = text.chars().any(char::is_alphabetic);
            (single_paragraph && not_numeric && text.chars().count() <= HEADING_MAX_CHARACTERS)
                .then(|| shape_typography(shape).map(|(size, bold)| (*index, size, bold)))?
        })
        .collect();
    if styled.is_empty() {
        return None;
    }
    let mut all_sizes: Vec<f64> = indices
        .iter()
        .filter_map(|index| match &slide.elements[*index] {
            SlideElement::Shape(shape) => shape_typography(shape).map(|value| value.0),
            _ => None,
        })
        .collect();
    if all_sizes.len() < 2 {
        return None;
    }
    all_sizes.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let median = all_sizes[(all_sizes.len() - 1) / 2].max(1.0);
    styled.sort_by(|a, b| {
        b.1.partial_cmp(&a.1)
            .unwrap_or(std::cmp::Ordering::Equal)
            .then_with(|| b.2.cmp(&a.2))
            .then(a.0.cmp(&b.0))
    });
    let (index, size, bold) = styled[0];
    (size >= median * HEADING_FONT_SIZE_RATIO || (bold && size >= median)).then_some(index)
}

fn attached_panel_label(slide: &Slide, candidate_index: usize, indices: &[usize]) -> bool {
    let SlideElement::Shape(candidate) = &slide.elements[candidate_index] else {
        return false;
    };
    let Some(text) = shape_text_plain(candidate) else {
        return false;
    };
    if candidate
        .text_body
        .as_ref()
        .is_none_or(|body| body.paragraphs.len() != 1)
        || text.chars().count() > HEADING_MAX_CHARACTERS
        || !text.chars().any(char::is_alphabetic)
    {
        return false;
    }
    let (x, y, width, height) = element_bounds(&slide.elements[candidate_index]);
    if width <= 0 || height <= 0 {
        return false;
    }
    indices.iter().copied().any(|other_index| {
        if other_index == candidate_index {
            return false;
        }
        let (other_x, other_y, other_width, other_height) =
            element_bounds(&slide.elements[other_index]);
        if other_width <= width || other_height <= height || y > other_y {
            return false;
        }
        let overlap_x = x
            .saturating_add(width)
            .min(other_x.saturating_add(other_width))
            .saturating_sub(x.max(other_x));
        let overlap_y = y
            .saturating_add(height)
            .min(other_y.saturating_add(other_height))
            .saturating_sub(y.max(other_y));
        overlap_x.saturating_mul(4) >= width.saturating_mul(3)
            && overlap_y.saturating_mul(5) >= height
            && width.saturating_mul(4) <= other_width.saturating_mul(3)
            && height.saturating_mul(2) <= other_height
    })
}

fn render_shape_heading_md(element: &SlideElement, out: &mut MarkdownWriter) {
    let SlideElement::Shape(shape) = element else {
        render_element_md(element, out);
        return;
    };
    let Some(body) = &shape.text_body else { return };
    let Some(paragraph) = body.paragraphs.first() else {
        return;
    };
    out.push_str("## ");
    render_runs_md(&paragraph.runs, paragraph, body, out);
    out.push_str("\n\n");
}

pub(crate) fn render_element_md(el: &SlideElement, out: &mut MarkdownWriter) {
    match el {
        SlideElement::Shape(s) => {
            let ph = s.placeholder_type.as_deref().unwrap_or("");
            // The slide-level # heading already used the title placeholder's
            // text — skip it here to avoid duplicating it inside the body.
            if ph == "title" || ph == "ctrTitle" {
                return;
            }
            // Drop auto-generated metadata placeholders (slide number, date,
            // footer, header). Their text is always a single token like "3" or
            // "2026-05-11" that's pure noise for an agent reading the content.
            if matches!(ph, "sldNum" | "dt" | "ftr" | "hdr") {
                return;
            }
            render_shape_md(s, out);
        }
        SlideElement::Table(t) => render_table_md(t, out),
        SlideElement::Chart(c) => render_chart_md(c, out),
        // Pictures / media / connectors carry no readable text; intentionally
        // dropped in the markdown projection. Use `pptx_get_pictures` or the
        // raw JSON path when you need to inspect them.
        SlideElement::Picture(_) | SlideElement::Media(_) => {}
    }
}

pub(crate) fn render_shape_md(s: &ShapeElement, out: &mut MarkdownWriter) {
    let Some(tb) = &s.text_body else { return };
    if tb.paragraphs.is_empty() {
        return;
    }
    // Body / subtitle placeholders inherit bullet formatting from the layout's
    // lstStyle (ECMA-376 §19.7.10) — treat `Bullet::Inherit` paragraphs there
    // as bulleted, mirroring what PowerPoint draws. Free text boxes default to
    // plain paragraphs.
    let ph = s.placeholder_type.as_deref().unwrap_or("");
    let inherit_means_bullet = matches!(ph, "body" | "subTitle" | "obj" | "tx" | "ftr" | "hdr");
    let leading_heading = inferred_leading_paragraph_heading(s);
    for (index, para) in tb.paragraphs.iter().enumerate() {
        let kind = paragraph_kind(&para.bullet, inherit_means_bullet);
        if index == 0 && leading_heading {
            out.push_str("## ");
            render_runs_md(&para.runs, para, tb, out);
            out.push_str("\n\n");
        } else {
            render_paragraph_md(para, tb, inherit_means_bullet, out);
        }
        // Plain text boxes often contain separate prose paragraphs. A blank
        // line is necessary for Markdown to preserve that authored boundary;
        // list items remain contiguous.
        if !(index == 0 && leading_heading)
            && matches!(kind, ParaKind::Plain)
            && index + 1 < tb.paragraphs.len()
        {
            out.push('\n');
        }
    }
    out.push('\n');
}

fn inferred_leading_paragraph_heading(shape: &ShapeElement) -> bool {
    let Some(body) = &shape.text_body else {
        return false;
    };
    if body.paragraphs.len() < 2 {
        return false;
    }
    let first = &body.paragraphs[0];
    let inherit_means_bullet = matches!(
        shape.placeholder_type.as_deref(),
        Some("body" | "subTitle" | "obj" | "tx")
    );
    if !matches!(
        paragraph_kind(&first.bullet, inherit_means_bullet),
        ParaKind::Plain
    ) {
        return false;
    }
    let text = paragraph_text_plain(first);
    if text.chars().count() > HEADING_MAX_CHARACTERS || !text.chars().any(char::is_alphabetic) {
        return false;
    }
    let first_bold = paragraph_is_bold(first, body);
    let following_bold = body.paragraphs[1..]
        .iter()
        .find(|paragraph| !paragraph_text_plain(paragraph).is_empty())
        .is_some_and(|paragraph| paragraph_is_bold(paragraph, body));
    if first_bold && !following_bold {
        return true;
    }
    let Some((first_size, _)) = paragraph_typography(first, body) else {
        return false;
    };
    let mut following_sizes: Vec<f64> = body.paragraphs[1..]
        .iter()
        .filter_map(|paragraph| paragraph_typography(paragraph, body).map(|value| value.0))
        .collect();
    if following_sizes.is_empty() {
        return false;
    }
    following_sizes.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let median = following_sizes[(following_sizes.len() - 1) / 2].max(1.0);
    first_size >= median * HEADING_FONT_SIZE_RATIO || (first_bold && first_size >= median)
}

fn shape_is_bold(shape: &ShapeElement) -> bool {
    shape.text_body.as_ref().is_some_and(|body| {
        body.paragraphs
            .iter()
            .any(|paragraph| paragraph_is_bold(paragraph, body))
    })
}

fn paragraph_is_bold(paragraph: &Paragraph, body: &TextBody) -> bool {
    paragraph.runs.iter().any(|run| {
        matches!(run, TextRun::Text(text) if !text.text.trim().is_empty()
            && text.bold.or(paragraph.def_bold).or(body.default_bold).unwrap_or(false))
    })
}

fn paragraph_typography(paragraph: &Paragraph, body: &TextBody) -> Option<(f64, bool)> {
    let mut max_size: Option<f64> = None;
    let mut bold = false;
    for run in &paragraph.runs {
        let TextRun::Text(run) = run else { continue };
        if run.text.trim().is_empty() {
            continue;
        }
        if let Some(size) = run
            .font_size
            .or(paragraph.def_font_size)
            .or(body.default_font_size)
        {
            max_size = Some(max_size.map_or(size, |current| current.max(size)));
        }
        bold |= run
            .bold
            .or(paragraph.def_bold)
            .or(body.default_bold)
            .unwrap_or(false);
    }
    max_size.map(|size| (size, bold))
}

fn paragraph_text_plain(paragraph: &Paragraph) -> String {
    paragraph
        .runs
        .iter()
        .filter_map(|run| match run {
            TextRun::Text(text) => Some(text.text.as_str()),
            _ => None,
        })
        .collect::<String>()
        .trim()
        .to_string()
}

#[derive(Clone, Copy)]
pub(crate) enum ParaKind {
    Plain,
    Bullet,
    Number,
}

pub(crate) fn paragraph_kind(b: &Bullet, inherit_means_bullet: bool) -> ParaKind {
    match b {
        Bullet::None => ParaKind::Plain,
        Bullet::Char { .. } => ParaKind::Bullet,
        // A picture bullet is still an unordered list item for markdown export.
        Bullet::Blip { .. } => ParaKind::Bullet,
        Bullet::AutoNum { .. } => ParaKind::Number,
        Bullet::Inherit => {
            if inherit_means_bullet {
                ParaKind::Bullet
            } else {
                ParaKind::Plain
            }
        }
    }
}

pub(crate) fn render_paragraph_md(
    para: &Paragraph,
    body: &TextBody,
    inherit_means_bullet: bool,
    out: &mut MarkdownWriter,
) {
    use std::fmt::Write as _;
    if !runs_have_visible_text(&para.runs) {
        out.push('\n');
        return;
    }
    for _ in 0..para.lvl {
        out.push_str("  ");
    }
    match paragraph_kind(&para.bullet, inherit_means_bullet) {
        ParaKind::Plain => {}
        ParaKind::Bullet => out.push_str("- "),
        // We deliberately emit `1.` for every numbered paragraph rather than
        // tracking the real counter — every markdown renderer auto-renumbers
        // sequential ordered-list items, so the visual output is correct and
        // we don't need to carry per-list state.
        ParaKind::Number => out.push_str("1. "),
    }
    render_runs_md(&para.runs, para, body, out);
    let _ = writeln!(out);
}

fn runs_have_visible_text(runs: &[TextRun]) -> bool {
    runs.iter().any(|run| match run {
        TextRun::Break => false,
        TextRun::Math { nodes, .. } => !nodes_to_text(nodes).trim().is_empty(),
        TextRun::Text(text) => !text.text.trim().is_empty(),
    })
}

pub(crate) fn render_runs_md(
    runs: &[TextRun],
    para: &Paragraph,
    body: &TextBody,
    out: &mut MarkdownWriter,
) {
    let mut index = 0;
    while index < runs.len() {
        match &runs[index] {
            // Intra-paragraph soft break (<a:br/>) → markdown hard line break
            // (two trailing spaces + newline).
            TextRun::Break => {
                out.push_str("  \n");
                index += 1;
            }
            // Equations have no faithful markdown form; emit their flattened text.
            TextRun::Math { nodes, .. } => {
                out.push_str(&nodes_to_text(nodes));
                index += 1;
            }
            TextRun::Text(t) => {
                let style = inline_style(t, para, body);
                let mut raw = t.text.clone();
                index += 1;
                // PowerPoint frequently splits a visually continuous phrase
                // into several runs for font/color metadata that Markdown
                // cannot represent. Coalesce adjacent runs when their
                // Markdown-visible style is identical so `**one phrase**`
                // does not become `**one** **phrase**`.
                while let Some(TextRun::Text(next)) = runs.get(index) {
                    if inline_style(next, para, body) != style {
                        break;
                    }
                    raw.push_str(&next.text);
                    index += 1;
                }
                render_text_span(&raw, style, out);
            }
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct InlineStyle<'a> {
    bold: bool,
    italic: bool,
    strikethrough: bool,
    hyperlink: Option<&'a str>,
}

fn inline_style<'a>(
    text: &'a TextRunData,
    paragraph: &Paragraph,
    body: &TextBody,
) -> InlineStyle<'a> {
    InlineStyle {
        bold: text
            .bold
            .or(paragraph.def_bold)
            .or(body.default_bold)
            .unwrap_or(false),
        italic: text
            .italic
            .or(paragraph.def_italic)
            .or(body.default_italic)
            .unwrap_or(false),
        strikethrough: text.strikethrough,
        hyperlink: text.hyperlink.as_deref(),
    }
}

fn render_text_span(raw: &str, style: InlineStyle<'_>, out: &mut MarkdownWriter) {
    // Empty / whitespace-only runs should not trigger formatting wrappers.
    if raw.chars().all(char::is_whitespace) {
        out.push_str(raw);
        return;
    }
    // Markdown strong/emphasis delimiters cannot contain edge whitespace.
    let leading_len = raw.len() - raw.trim_start().len();
    let trail_start = raw.trim_end().len();
    let leading = &raw[..leading_len];
    let trailing = &raw[trail_start..];
    let trimmed = &raw[leading_len..trail_start];
    out.push_str(leading);
    if style.strikethrough {
        out.push_str("~~");
    }
    if style.italic {
        out.push('*');
    }
    if style.bold {
        out.push_str("**");
    }
    if style.hyperlink.is_some() {
        out.push('[');
    }
    write_escaped_inline_md(trimmed, out);
    if let Some(url) = style.hyperlink {
        out.push_str("](");
        out.push_str(url);
        out.push(')');
    }
    if style.bold {
        out.push_str("**");
    }
    if style.italic {
        out.push('*');
    }
    if style.strikethrough {
        out.push_str("~~");
    }
    out.push_str(trailing);
}

/// Escape the markdown inline metacharacters that would otherwise be parsed as
/// formatting. We deliberately don't escape every potential metachar — pptx
/// body text contains so much punctuation that aggressive escaping makes the
/// output noisier than the structure it's trying to expose. Pipe is handled
/// separately in `render_table_cell_md` since it only matters inside tables.
fn write_escaped_inline_md(value: &str, out: &mut MarkdownWriter) {
    for character in value.chars() {
        if matches!(character, '\\' | '*' | '_' | '`') {
            out.push('\\');
        }
        out.push(character);
    }
}

pub(crate) fn render_table_md(t: &TableElement, out: &mut MarkdownWriter) {
    use std::fmt::Write as _;
    if t.rows.is_empty() {
        return;
    }
    let cols = t.rows[0].cells.len();
    if cols == 0 {
        return;
    }
    let header_cells: Vec<String> = t.rows[0].cells.iter().map(render_table_cell_md).collect();
    let _ = writeln!(out, "| {} |", header_cells.join(" | "));
    let sep: Vec<&str> = (0..cols).map(|_| "---").collect();
    let _ = writeln!(out, "| {} |", sep.join(" | "));
    for row in t.rows.iter().skip(1) {
        let cells: Vec<String> = row.cells.iter().map(render_table_cell_md).collect();
        let _ = writeln!(out, "| {} |", cells.join(" | "));
    }
    out.push('\n');
}

pub(crate) fn render_table_cell_md(cell: &TableCell) -> String {
    // Continuation cells of a merge carry no content — leave empty so the row
    // alignment stays intact.
    if cell.h_merge || cell.v_merge {
        return String::new();
    }
    let Some(tb) = &cell.text_body else {
        return String::new();
    };
    let mut buf = String::new();
    for (i, para) in tb.paragraphs.iter().enumerate() {
        if i > 0 {
            buf.push_str("<br>");
        }
        for run in &para.runs {
            if let TextRun::Text(t) = run {
                buf.push_str(&t.text);
            }
        }
    }
    buf.trim().replace('|', "\\|")
}

pub(crate) fn render_chart_md(c: &ChartElement, out: &mut MarkdownWriter) {
    use std::fmt::Write as _;
    let chart = &c.chart;
    let title = chart.title.as_deref().unwrap_or("(untitled)");
    let _ = writeln!(out, "**Chart ({}): {}**\n", chart.chart_type, title);
    if !chart.categories.is_empty() {
        let _ = writeln!(out, "- Categories: {}", chart.categories.join(", "));
    }
    for s in &chart.series {
        let values: Vec<String> = s
            .values
            .iter()
            .map(|v| match v {
                Some(n) => format!("{n}"),
                None => "—".to_string(),
            })
            .collect();
        let _ = writeln!(out, "- {}: {}", s.name, values.join(", "));
    }
    out.push('\n');
}

/// Append slide comments to a presentation-level review section. Keeping
/// review metadata out of the slide narrative prevents comments from breaking
/// a heading, list, table, or spatially-derived reading block.
pub(crate) fn render_slide_comments_md(
    slide: &Slide,
    slide_width: i64,
    slide_height: i64,
    appendix: &mut MarkdownWriter,
    has_comments: &mut bool,
) {
    use std::fmt::Write as _;
    if slide.comments.is_empty() {
        return;
    }
    if !*has_comments {
        appendix.push_str("\n## Review comments\n\n");
        *has_comments = true;
    }
    let title = project_slide(slide, slide_width, slide_height)
        .title
        .map(|(_, title)| title);
    for comment in &slide.comments {
        let target = comment_target(slide, comment)
            .or_else(|| title.clone())
            .unwrap_or_else(|| format!("Slide {}", slide.slide_number));
        let author = comment.author.as_deref().unwrap_or("(unknown)");
        let status = comment
            .status
            .as_deref()
            .filter(|status| *status != "active")
            .map(|status| format!(" [{status}]"))
            .unwrap_or_default();
        let _ = writeln!(
            appendix,
            "### Slide {} — {}\n\n**{}{}**\n",
            slide.slide_number,
            escape_heading_text(&target),
            author,
            status
        );
        write_blockquote(comment.text.trim(), "> ", appendix);
        for reply in &comment.replies {
            let reply_author = reply.author.as_deref().unwrap_or("(unknown)");
            let reply_status = reply
                .status
                .as_deref()
                .filter(|status| *status != "active")
                .map(|status| format!(" [{status}]"))
                .unwrap_or_default();
            let _ = writeln!(appendix, ">> **{}{}**", reply_author, reply_status);
            write_blockquote(reply.text.trim(), ">> ", appendix);
        }
        appendix.push('\n');
    }
}

fn write_blockquote(value: &str, prefix: &str, out: &mut MarkdownWriter) {
    use std::fmt::Write as _;
    if value.is_empty() {
        let _ = writeln!(out, "{prefix}");
        return;
    }
    for line in value.lines() {
        let _ = writeln!(out, "{prefix}{line}");
    }
}

fn escape_heading_text(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('*', "\\*")
        .replace('_', "\\_")
        .replace('`', "\\`")
        .replace('#', "\\#")
}

fn comment_target(slide: &Slide, comment: &PptxComment) -> Option<String> {
    for anchor in &comment.anchors {
        let element_id = match anchor {
            PptxCommentAnchor::DrawingElement { element_id, .. }
            | PptxCommentAnchor::TextRange { element_id, .. } => element_id.as_deref(),
            PptxCommentAnchor::Slide | PptxCommentAnchor::Unknown => None,
        };
        if let Some(id) = element_id {
            if let Some(label) = slide.elements.iter().find_map(|element| {
                (element_id_of(element) == Some(id))
                    .then(|| element_label(element))
                    .flatten()
            }) {
                return Some(label);
            }
        }
    }
    let (x, y) = (comment.x?, comment.y?);
    slide
        .elements
        .iter()
        .filter_map(|element| {
            // A classic comment position is only a point on the slide, not an
            // explicit element relationship. Restrict proximity labels to
            // readable content so decorative lines/pictures do not become
            // misleading targets.
            let label = element_content_label(element)?;
            let (left, top, width, height) = element_bounds(element);
            let right = left.saturating_add(width);
            let bottom = top.saturating_add(height);
            (x >= left && x <= right && y >= top && y <= bottom)
                .then_some((width.saturating_mul(height), label))
        })
        .min_by_key(|(area, _)| *area)
        .map(|(_, label)| label)
}

fn element_content_label(element: &SlideElement) -> Option<String> {
    match element {
        SlideElement::Shape(shape) => shape_text_plain(shape),
        SlideElement::Table(_) => Some("Table".to_string()),
        SlideElement::Chart(chart) => Some(
            chart
                .chart
                .title
                .clone()
                .unwrap_or_else(|| "Chart".to_string()),
        ),
        SlideElement::Picture(_) | SlideElement::Media(_) => None,
    }
    .map(compact_label)
    .filter(|label| !label.is_empty())
}

fn element_id_of(element: &SlideElement) -> Option<&str> {
    match element {
        SlideElement::Shape(value) => value.id.as_deref(),
        SlideElement::Picture(value) => value.id.as_deref(),
        SlideElement::Table(value) => value.id.as_deref(),
        SlideElement::Chart(value) => value.id.as_deref(),
        SlideElement::Media(value) => value.id.as_deref(),
    }
}

fn element_label(element: &SlideElement) -> Option<String> {
    match element {
        SlideElement::Shape(shape) => shape_text_plain(shape).or_else(|| shape.name.clone()),
        SlideElement::Table(_) => Some("Table".to_string()),
        SlideElement::Chart(chart) => Some(
            chart
                .chart
                .title
                .clone()
                .unwrap_or_else(|| "Chart".to_string()),
        ),
        SlideElement::Picture(_) => Some("Picture".to_string()),
        SlideElement::Media(media) => Some(media.media_kind.clone()),
    }
    .map(compact_label)
    .filter(|label| !label.is_empty())
}

fn compact_label(label: String) -> String {
    let mut chars = label.trim().chars();
    let compact: String = chars.by_ref().take(80).collect();
    if chars.next().is_some() {
        format!("{compact}…")
    } else {
        compact
    }
}

fn element_bounds(element: &SlideElement) -> (i64, i64, i64, i64) {
    match element {
        SlideElement::Shape(value) => (value.x, value.y, value.width, value.height),
        SlideElement::Picture(value) => (value.x, value.y, value.width, value.height),
        SlideElement::Table(value) => (value.x, value.y, value.width, value.height),
        SlideElement::Chart(value) => (value.x, value.y, value.width, value.height),
        SlideElement::Media(value) => (value.x, value.y, value.width, value.height),
    }
}

/// Materialized-model oracle retained for compatibility/degraded paths and
/// sequential-output equivalence tests.
pub(crate) fn render_presentation_md(pres: &Presentation) -> String {
    let mut out = MarkdownWriter::new(u64::MAX);
    let mut appendix = MarkdownWriter::sharing_budget(&out);
    let mut has_comments = false;
    for (i, slide) in pres.slides.iter().enumerate() {
        if i > 0 {
            out.push_str("\n---\n\n");
        }
        render_slide_md(slide, pres.slide_width, pres.slide_height, &mut out);
        render_slide_comments_md(
            slide,
            pres.slide_width,
            pres.slide_height,
            &mut appendix,
            &mut has_comments,
        );
    }
    out.append_precounted(appendix.into_string());
    out.into_string()
}

#[cfg(test)]
mod tests {
    use super::MarkdownWriter;

    #[test]
    fn writer_counts_utf8_bytes_and_stops_retaining_at_the_first_crossing() {
        let mut output = MarkdownWriter::new(3);
        output.push_str("é");
        output.push_str("é");
        output.push_str("ignored after crossing");

        assert_eq!(output.observed(), 4);
        assert_eq!(output.into_string(), "é");
    }
}
