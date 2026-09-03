//! Semantic layout projection for PPTX markdown.
//!
//! PresentationML's shape-tree order is paint/navigation order (ECMA-376 Part
//! 1 §19.3.1.45), not prose reading order. This module keeps that source order
//! as a deterministic tie-breaker while deriving Markdown order from authored
//! groups, containing panel shapes, and non-overlapping spatial partitions.

use crate::types::*;
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

const PANEL_CONTAINMENT_TOLERANCE_RATIO: f64 = 0.015;
const PANEL_MAX_SLIDE_AREA_RATIO: f64 = 0.9;
const PANEL_COMPONENT_OVERLAP_RATIO: f64 = 0.08;
const ATTACHED_CONTENT_OVERLAP_RATIO: f64 = 0.15;
const DUPLICATE_CONTENT_OVERLAP_RATIO: f64 = 0.8;
const INDEPENDENT_COLUMN_MIN_SLIDE_HEIGHT_RATIO: f64 = 0.35;
const INDEPENDENT_COLUMN_MIN_SLIDE_WIDTH_RATIO: f64 = 0.15;
const INDEPENDENT_COLUMN_MIN_VERTICAL_OVERLAP_RATIO: f64 = 0.5;
const GRID_ROW_OVERLAP_RATIO: f64 = 0.5;
const GRID_ALIGNMENT_TOLERANCE_RATIO: f64 = 0.05;
const PLACEHOLDER_TITLE_REGION_NUMERATOR: i64 = 9;
const PLACEHOLDER_TITLE_REGION_DENOMINATOR: i64 = 20;
const INFERRED_TITLE_REGION_NUMERATOR: i64 = 3;
const INFERRED_TITLE_REGION_DENOMINATOR: i64 = 10;
const INFERRED_TITLE_MIN_WIDTH_DENOMINATOR: i64 = 4;
const HEADING_MAX_CHARACTERS: usize = 140;
const TITLE_FONT_SIZE_RATIO: f64 = 1.25;
const BOLD_TITLE_FONT_SIZE_RATIO: f64 = 1.1;
const LOW_PLACEHOLDER_TITLE_FONT_SIZE_RATIO: f64 = 1.5;

#[derive(Clone, Copy, Debug, PartialEq)]
struct Rect {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
}

impl Rect {
    fn right(self) -> f64 {
        self.x + self.width.max(0.0)
    }

    fn bottom(self) -> f64 {
        self.y + self.height.max(0.0)
    }

    fn area(self) -> f64 {
        self.width.max(0.0) * self.height.max(0.0)
    }

    fn intersection_area(self, other: Self) -> f64 {
        let width = self.right().min(other.right()) - self.x.max(other.x);
        let height = self.bottom().min(other.bottom()) - self.y.max(other.y);
        width.max(0.0) * height.max(0.0)
    }

    fn overlap_ratio(self, other: Self) -> f64 {
        self.intersection_area(other) / self.area().min(other.area()).max(1.0)
    }

    fn union(self, other: Self) -> Self {
        let x = self.x.min(other.x);
        let y = self.y.min(other.y);
        let right = self.right().max(other.right());
        let bottom = self.bottom().max(other.bottom());
        Self {
            x,
            y,
            width: right - x,
            height: bottom - y,
        }
    }

    fn contains(self, other: Self) -> bool {
        // A small relative allowance absorbs rounding introduced by nested
        // group transforms without letting nearby cards bleed together.
        let tolerance = self.width.min(self.height).max(1.0) * PANEL_CONTAINMENT_TOLERANCE_RATIO;
        other.x >= self.x - tolerance
            && other.y >= self.y - tolerance
            && other.right() <= self.right() + tolerance
            && other.bottom() <= self.bottom() + tolerance
    }
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) struct SemanticBlock {
    pub(crate) element_indices: Vec<usize>,
    /// True for an authored group or an inferred containing panel. Markdown
    /// uses this to infer a local heading only where a real relationship exists.
    pub(crate) related: bool,
    /// A thematic break before this block preserves an independent spatial
    /// region after the two-dimensional slide is linearized.
    pub(crate) starts_new_region: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) struct SemanticSlide {
    pub(crate) title: Option<(usize, String)>,
    pub(crate) blocks: Vec<SemanticBlock>,
}

#[derive(Clone, Debug)]
struct LayoutBlock {
    rect: Rect,
    source_order: usize,
    element_indices: Vec<usize>,
    related: bool,
    starts_new_region: bool,
}

/// Resource-governance ceiling for recursive whitespace partitioning. A slide
/// with more nested separators still gets deterministic geometry order, but a
/// hostile shape tree cannot turn Markdown projection into unbounded recursion.
const MAX_SPATIAL_PARTITION_DEPTH: usize = 64;

pub(crate) fn project_slide(slide: &Slide, slide_width: i64, slide_height: i64) -> SemanticSlide {
    let title = title_candidate(slide, slide_width, slide_height);
    let title_index = title.as_ref().map(|(index, _)| *index);
    let candidates: Vec<usize> = slide
        .elements
        .iter()
        .enumerate()
        .filter_map(|(index, element)| {
            (Some(index) != title_index && element_has_content(element)).then_some(index)
        })
        .collect();
    let visible = deduplicate_visible_content(slide, candidates);

    let mut claimed = HashSet::new();
    let mut blocks = Vec::new();

    // Explicit PowerPoint grouping is the strongest relationship signal.
    for group in &slide.semantic_groups {
        let members: Vec<usize> = visible
            .iter()
            .copied()
            .filter(|index| *index >= group.start && *index < group.end)
            .collect();
        if members.is_empty() {
            continue;
        }
        claimed.extend(members.iter().copied());
        blocks.push(make_block(slide, members, true));
    }

    // A common ungrouped authoring pattern is a set of overlapping filled or
    // stroked shapes (for example a Venn diagram) behind multiple text boxes.
    // Treat the connected underlays as one semantic region, then associate
    // content with the smallest containing shape. Full-slide backgrounds are
    // deliberately excluded: they describe the slide, not a local block.
    let slide_area = (slide_width.max(1) as f64) * (slide_height.max(1) as f64);
    let panels: Vec<(usize, Rect)> = slide
        .elements
        .iter()
        .enumerate()
        .filter(|(_, element)| !element_has_content(element) && is_panel_shape(element))
        .filter_map(|(index, element)| {
            element_rect(element)
                .filter(|rect| rect.area() < slide_area * PANEL_MAX_SLIDE_AREA_RATIO)
                .map(|rect| (index, rect))
        })
        .collect();
    let mut panel_memberships: HashMap<usize, Vec<usize>> = HashMap::new();
    for index in visible
        .iter()
        .copied()
        .filter(|index| !claimed.contains(index))
    {
        let Some(content_rect) = element_rect(&slide.elements[index]) else {
            continue;
        };
        let panel = panels
            .iter()
            .filter(|(panel_index, rect)| *panel_index < index && rect.contains(content_rect))
            .map(|(panel_index, rect)| (*panel_index, rect.area()))
            .min_by(|a, b| a.1.partial_cmp(&b.1).unwrap_or(Ordering::Equal))
            .map(|(panel_index, _)| panel_index);
        if let Some(panel_index) = panel {
            panel_memberships
                .entry(panel_index)
                .or_default()
                .push(index);
        }
    }
    // Only panels that actually contain content may connect regions. This
    // prevents a large decorative outline or flourish from acting as a bridge
    // between otherwise independent cards merely because its bounding box
    // overlaps them.
    let active_panels: Vec<(usize, Rect)> = panels
        .iter()
        .copied()
        .filter(|(index, _)| panel_memberships.contains_key(index))
        .collect();
    for panel_component in connected_components(&active_panels, PANEL_COMPONENT_OVERLAP_RATIO) {
        let members: Vec<usize> = panel_component
            .iter()
            .flat_map(|panel| panel_memberships.get(panel).into_iter().flatten().copied())
            .collect();
        if members.len() < 2 {
            continue;
        }
        claimed.extend(members.iter().copied());
        blocks.push(make_block(slide, members, true));
    }

    // Label badges and their content-bearing cards are often authored as two
    // ungrouped shapes that intentionally overlap. A substantial overlap with
    // at least one styled panel is a stronger relationship signal than mere
    // proximity, so merge these attachments before spatial ordering.
    let remaining: Vec<(usize, Rect)> = visible
        .iter()
        .copied()
        .filter(|index| !claimed.contains(index))
        .filter_map(|index| element_rect(&slide.elements[index]).map(|rect| (index, rect)))
        .collect();
    for members in connected_content_components(slide, &remaining) {
        if members.len() < 2 {
            continue;
        }
        claimed.extend(members.iter().copied());
        blocks.push(make_block(slide, members, true));
    }

    for index in visible.into_iter().filter(|index| !claimed.contains(index)) {
        blocks.push(make_block(slide, vec![index], false));
    }

    let blocks = spatial_order(
        blocks,
        slide_width.max(1) as f64,
        slide_height.max(1) as f64,
    )
    .into_iter()
    .map(|mut block| {
        if block.element_indices.len() > 1 {
            block.element_indices.sort_by(|a, b| {
                compare_rects(
                    element_rect(&slide.elements[*a]),
                    element_rect(&slide.elements[*b]),
                    *a,
                    *b,
                )
            });
        }
        SemanticBlock {
            element_indices: block.element_indices,
            related: block.related,
            starts_new_region: block.starts_new_region,
        }
    })
    .collect();

    SemanticSlide { title, blocks }
}

fn make_block(slide: &Slide, element_indices: Vec<usize>, related: bool) -> LayoutBlock {
    let source_order = *element_indices.iter().min().unwrap_or(&usize::MAX);
    let rect = element_indices
        .iter()
        .filter_map(|index| element_rect(&slide.elements[*index]))
        .reduce(Rect::union)
        .unwrap_or(Rect {
            x: 0.0,
            y: 0.0,
            width: 0.0,
            height: 0.0,
        });
    LayoutBlock {
        rect,
        source_order,
        element_indices,
        related,
        starts_new_region: false,
    }
}

fn spatial_order(
    blocks: Vec<LayoutBlock>,
    slide_width: f64,
    slide_height: f64,
) -> Vec<LayoutBlock> {
    spatial_order_at_depth(blocks, slide_width, slide_height, 0)
}

fn spatial_order_at_depth(
    blocks: Vec<LayoutBlock>,
    slide_width: f64,
    slide_height: f64,
    depth: usize,
) -> Vec<LayoutBlock> {
    if blocks.len() <= 1 {
        return blocks;
    }
    if depth >= MAX_SPATIAL_PARTITION_DEPTH {
        let mut fallback = blocks;
        fallback.sort_by(|a, b| {
            compare_rects(Some(a.rect), Some(b.rect), a.source_order, b.source_order)
        });
        return fallback;
    }
    let horizontal = best_split(&blocks, Axis::Horizontal, slide_height);
    let vertical = best_split(&blocks, Axis::Vertical, slide_width);
    let selected = match (horizontal, vertical) {
        (Some(h), Some(v)) => {
            // A real repeated grid reads row-major. Independent tall columns
            // read one complete region at a time, even when small horizontal
            // gaps inside the columns happen to line up.
            if looks_like_repeated_grid(&blocks, slide_width) {
                h
            } else if independent_column_split(&blocks, &v, slide_width, slide_height) {
                v
            } else if h.score >= v.score {
                h
            } else {
                v
            }
        }
        (Some(split), None) | (None, Some(split)) => split,
        (None, None) => {
            let mut fallback = blocks;
            fallback.sort_by(|a, b| {
                compare_rects(Some(a.rect), Some(b.rect), a.source_order, b.source_order)
            });
            return fallback;
        }
    };

    let starts_independent_region = matches!(selected.axis, Axis::Vertical)
        && independent_column_split(&blocks, &selected, slide_width, slide_height);
    let mut first = Vec::new();
    let mut second = Vec::new();
    for block in blocks {
        if selected.first_orders.contains(&block.source_order) {
            first.push(block);
        } else {
            second.push(block);
        }
    }
    let mut result = spatial_order_at_depth(first, slide_width, slide_height, depth + 1);
    let mut ordered_second = spatial_order_at_depth(second, slide_width, slide_height, depth + 1);
    if starts_independent_region {
        if let Some(block) = ordered_second.first_mut() {
            block.starts_new_region = true;
        }
    }
    result.extend(ordered_second);
    result
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Axis {
    Horizontal,
    Vertical,
}

struct Split {
    axis: Axis,
    first_orders: HashSet<usize>,
    score: f64,
}

fn best_split(blocks: &[LayoutBlock], axis: Axis, extent: f64) -> Option<Split> {
    let mut ordered: Vec<&LayoutBlock> = blocks.iter().collect();
    ordered.sort_by(|a, b| {
        let av = match axis {
            Axis::Horizontal => a.rect.y,
            Axis::Vertical => a.rect.x,
        };
        let bv = match axis {
            Axis::Horizontal => b.rect.y,
            Axis::Vertical => b.rect.x,
        };
        av.partial_cmp(&bv)
            .unwrap_or(Ordering::Equal)
            .then(a.source_order.cmp(&b.source_order))
    });
    let mut best: Option<Split> = None;
    for cut in 1..ordered.len() {
        let first_end = ordered[..cut]
            .iter()
            .map(|block| match axis {
                Axis::Horizontal => block.rect.bottom(),
                Axis::Vertical => block.rect.right(),
            })
            .fold(f64::NEG_INFINITY, f64::max);
        let second_start = ordered[cut..]
            .iter()
            .map(|block| match axis {
                Axis::Horizontal => block.rect.y,
                Axis::Vertical => block.rect.x,
            })
            .fold(f64::INFINITY, f64::min);
        let gap = second_start - first_end;
        if gap <= 0.0 {
            continue;
        }
        let score = gap / extent.max(1.0);
        if best.as_ref().is_none_or(|current| score > current.score) {
            best = Some(Split {
                axis,
                first_orders: ordered[..cut]
                    .iter()
                    .map(|block| block.source_order)
                    .collect(),
                score,
            });
        }
    }
    best
}

fn independent_column_split(
    blocks: &[LayoutBlock],
    split: &Split,
    slide_width: f64,
    slide_height: f64,
) -> bool {
    if split.axis != Axis::Vertical {
        return false;
    }
    let first = blocks
        .iter()
        .filter(|block| split.first_orders.contains(&block.source_order))
        .map(|block| block.rect)
        .reduce(Rect::union);
    let second = blocks
        .iter()
        .filter(|block| !split.first_orders.contains(&block.source_order))
        .map(|block| block.rect)
        .reduce(Rect::union);
    let (Some(first), Some(second)) = (first, second) else {
        return false;
    };
    if first.width < slide_width * INDEPENDENT_COLUMN_MIN_SLIDE_WIDTH_RATIO
        || second.width < slide_width * INDEPENDENT_COLUMN_MIN_SLIDE_WIDTH_RATIO
        || first.height < slide_height * INDEPENDENT_COLUMN_MIN_SLIDE_HEIGHT_RATIO
        || second.height < slide_height * INDEPENDENT_COLUMN_MIN_SLIDE_HEIGHT_RATIO
    {
        return false;
    }
    let vertical_overlap = first.bottom().min(second.bottom()) - first.y.max(second.y);
    vertical_overlap.max(0.0) / first.height.min(second.height).max(1.0)
        >= INDEPENDENT_COLUMN_MIN_VERTICAL_OVERLAP_RATIO
}

fn looks_like_repeated_grid(blocks: &[LayoutBlock], slide_width: f64) -> bool {
    if blocks.len() < 4 {
        return false;
    }
    let mut ordered: Vec<&LayoutBlock> = blocks.iter().collect();
    ordered.sort_by(|a, b| {
        a.rect
            .y
            .partial_cmp(&b.rect.y)
            .unwrap_or(Ordering::Equal)
            .then_with(|| a.rect.x.partial_cmp(&b.rect.x).unwrap_or(Ordering::Equal))
    });
    let mut rows: Vec<Vec<&LayoutBlock>> = Vec::new();
    for block in ordered {
        let row = rows.iter_mut().find(|row| {
            let row_rect = row
                .iter()
                .map(|member| member.rect)
                .reduce(Rect::union)
                .unwrap_or(block.rect);
            let overlap = row_rect.bottom().min(block.rect.bottom()) - row_rect.y.max(block.rect.y);
            overlap.max(0.0) / row_rect.height.min(block.rect.height).max(1.0)
                >= GRID_ROW_OVERLAP_RATIO
        });
        if let Some(row) = row {
            row.push(block);
        } else {
            rows.push(vec![block]);
        }
    }
    if rows.len() < 2 {
        return false;
    }
    let columns = rows[0].len();
    if columns < 2 || rows.iter().any(|row| row.len() != columns) {
        return false;
    }
    for row in &mut rows {
        row.sort_by(|a, b| a.rect.x.partial_cmp(&b.rect.x).unwrap_or(Ordering::Equal));
    }
    let tolerance = slide_width.max(1.0) * GRID_ALIGNMENT_TOLERANCE_RATIO;
    (0..columns).all(|column| {
        let anchor = rows[0][column].rect;
        rows[1..].iter().all(|row| {
            (row[column].rect.x - anchor.x).abs() <= tolerance
                && (row[column].rect.right() - anchor.right()).abs() <= tolerance
        })
    })
}

fn deduplicate_visible_content(slide: &Slide, candidates: Vec<usize>) -> Vec<usize> {
    let mut visible = Vec::new();
    for index in candidates {
        let duplicate = visible.iter().position(|existing| {
            semantic_duplicate(&slide.elements[*existing], &slide.elements[index])
        });
        if let Some(position) = duplicate {
            // Keep the foreground copy while preserving the first occurrence's
            // reading position.
            visible[position] = index;
        } else {
            visible.push(index);
        }
    }
    visible
}

fn semantic_duplicate(first: &SlideElement, second: &SlideElement) -> bool {
    let (SlideElement::Shape(first), SlideElement::Shape(second)) = (first, second) else {
        return false;
    };
    let Some(first_text) = shape_text(first) else {
        return false;
    };
    let Some(second_text) = shape_text(second) else {
        return false;
    };
    if normalize_text(&first_text) != normalize_text(&second_text) {
        return false;
    }
    let (Some(first_rect), Some(second_rect)) = (shape_rect(first), shape_rect(second)) else {
        return false;
    };
    first_rect.overlap_ratio(second_rect) >= DUPLICATE_CONTENT_OVERLAP_RATIO
}

fn normalize_text(value: &str) -> String {
    value.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn connected_components(items: &[(usize, Rect)], overlap_threshold: f64) -> Vec<Vec<usize>> {
    let mut visited = vec![false; items.len()];
    let mut components = Vec::new();
    for start in 0..items.len() {
        if visited[start] {
            continue;
        }
        visited[start] = true;
        let mut stack = vec![start];
        let mut component = Vec::new();
        while let Some(current) = stack.pop() {
            component.push(items[current].0);
            for candidate in 0..items.len() {
                if !visited[candidate]
                    && items[current].1.overlap_ratio(items[candidate].1) >= overlap_threshold
                {
                    visited[candidate] = true;
                    stack.push(candidate);
                }
            }
        }
        components.push(component);
    }
    components
}

fn connected_content_components(slide: &Slide, items: &[(usize, Rect)]) -> Vec<Vec<usize>> {
    let mut visited = vec![false; items.len()];
    let mut components = Vec::new();
    for start in 0..items.len() {
        if visited[start] {
            continue;
        }
        visited[start] = true;
        let mut stack = vec![start];
        let mut component = Vec::new();
        while let Some(current) = stack.pop() {
            component.push(items[current].0);
            for candidate in 0..items.len() {
                if visited[candidate] {
                    continue;
                }
                let related = (is_panel_shape(&slide.elements[items[current].0])
                    || is_panel_shape(&slide.elements[items[candidate].0]))
                    && items[current].1.overlap_ratio(items[candidate].1)
                        >= ATTACHED_CONTENT_OVERLAP_RATIO;
                if related {
                    visited[candidate] = true;
                    stack.push(candidate);
                }
            }
        }
        components.push(component);
    }
    components
}

fn compare_rects(a: Option<Rect>, b: Option<Rect>, a_order: usize, b_order: usize) -> Ordering {
    match (a, b) {
        (Some(a), Some(b)) => {
            a.y.partial_cmp(&b.y)
                .unwrap_or(Ordering::Equal)
                .then_with(|| a.x.partial_cmp(&b.x).unwrap_or(Ordering::Equal))
                .then(a_order.cmp(&b_order))
        }
        _ => a_order.cmp(&b_order),
    }
}

fn title_candidate(slide: &Slide, slide_width: i64, slide_height: i64) -> Option<(usize, String)> {
    // Placeholder type is the strongest semantic hint, but real decks
    // sometimes reuse a title placeholder as a bottom attribution/caption.
    // Require it to occupy the upper title region or be typographically
    // dominant before promoting it to `#`.
    for (index, element) in slide.elements.iter().enumerate() {
        if let SlideElement::Shape(shape) = element {
            if matches!(
                shape.placeholder_type.as_deref(),
                Some("title" | "ctrTitle")
            ) && (shape.y
                < slide_height.saturating_mul(PLACEHOLDER_TITLE_REGION_NUMERATOR)
                    / PLACEHOLDER_TITLE_REGION_DENOMINATOR
                || low_placeholder_is_typographically_dominant(slide, shape))
            {
                if let Some(text) = shape_text(shape) {
                    return Some((index, text));
                }
            }
        }
    }

    let mut candidates = Vec::new();
    let mut body_sizes = Vec::new();
    for (index, element) in slide.elements.iter().enumerate() {
        let SlideElement::Shape(shape) = element else {
            continue;
        };
        if matches!(
            shape.placeholder_type.as_deref(),
            Some("sldNum" | "dt" | "ftr" | "hdr")
        ) {
            continue;
        }
        let Some(text) = shape_text(shape) else {
            continue;
        };
        let Some((font_size, bold)) = shape_typography(shape) else {
            continue;
        };
        body_sizes.push(font_size);
        let paragraph_count = shape
            .text_body
            .as_ref()
            .map_or(0, |body| body.paragraphs.len());
        let top_region = shape.y
            < slide_height.saturating_mul(INFERRED_TITLE_REGION_NUMERATOR)
                / INFERRED_TITLE_REGION_DENOMINATOR;
        let substantial_width =
            shape.width >= slide_width.saturating_div(INFERRED_TITLE_MIN_WIDTH_DENOMINATOR);
        let short = text.chars().count() <= HEADING_MAX_CHARACTERS && paragraph_count <= 2;
        let numeric_only = text.chars().all(|character| !character.is_alphabetic());
        if top_region && substantial_width && short && !numeric_only {
            candidates.push((index, text, font_size, bold, shape.y));
        }
    }
    if candidates.is_empty() || body_sizes.len() < 2 {
        return None;
    }
    body_sizes.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    let median = body_sizes[(body_sizes.len() - 1) / 2].max(1.0);
    candidates
        .into_iter()
        .filter(|(_, _, size, bold, _)| {
            *size >= median * TITLE_FONT_SIZE_RATIO
                || (*bold && *size >= median * BOLD_TITLE_FONT_SIZE_RATIO)
        })
        .max_by(|a, b| {
            a.2.partial_cmp(&b.2)
                .unwrap_or(Ordering::Equal)
                .then_with(|| b.4.cmp(&a.4))
        })
        .map(|(index, text, _, _, _)| (index, text))
}

fn low_placeholder_is_typographically_dominant(slide: &Slide, title: &ShapeElement) -> bool {
    let Some((title_size, _)) = shape_typography(title) else {
        return false;
    };
    let mut sizes: Vec<f64> = slide
        .elements
        .iter()
        .filter_map(|element| match element {
            SlideElement::Shape(shape)
                if !matches!(
                    shape.placeholder_type.as_deref(),
                    Some("sldNum" | "dt" | "ftr" | "hdr")
                ) =>
            {
                shape_typography(shape).map(|value| value.0)
            }
            _ => None,
        })
        .collect();
    if sizes.len() < 2 {
        return false;
    }
    sizes.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    let median = sizes[(sizes.len() - 1) / 2].max(1.0);
    title_size >= median * LOW_PLACEHOLDER_TITLE_FONT_SIZE_RATIO
}

pub(crate) fn shape_typography(shape: &ShapeElement) -> Option<(f64, bool)> {
    let body = shape.text_body.as_ref()?;
    let mut max_size: Option<f64> = None;
    let mut bold = false;
    for paragraph in &body.paragraphs {
        for run in &paragraph.runs {
            let TextRun::Text(run) = run else { continue };
            if run.text.trim().is_empty() {
                continue;
            }
            let size = run
                .font_size
                .or(paragraph.def_font_size)
                .or(body.default_font_size);
            if let Some(size) = size {
                max_size = Some(max_size.map_or(size, |current| current.max(size)));
            }
            bold |= run
                .bold
                .or(paragraph.def_bold)
                .or(body.default_bold)
                .unwrap_or(false);
        }
    }
    max_size.map(|size| (size, bold))
}

fn shape_text(shape: &ShapeElement) -> Option<String> {
    let body = shape.text_body.as_ref()?;
    let text = body
        .paragraphs
        .iter()
        .map(|paragraph| {
            paragraph
                .runs
                .iter()
                .filter_map(|run| match run {
                    TextRun::Text(text) => Some(text.text.as_str()),
                    _ => None,
                })
                .collect::<String>()
        })
        .collect::<Vec<_>>()
        .join(" ");
    let text = text.trim();
    (!text.is_empty()).then(|| text.to_string())
}

fn element_has_content(element: &SlideElement) -> bool {
    match element {
        SlideElement::Shape(shape) => {
            !matches!(shape.placeholder_type.as_deref(), Some("sldNum" | "dt" | "ftr" | "hdr"))
                && shape_text(shape).is_some()
        }
        SlideElement::Table(table) => table.rows.iter().any(|row| {
            row.cells.iter().any(|cell| {
                cell.text_body.as_ref().is_some_and(|body| {
                    body.paragraphs.iter().any(|paragraph| {
                        paragraph.runs.iter().any(|run| matches!(run, TextRun::Text(text) if !text.text.trim().is_empty()))
                    })
                })
            })
        }),
        SlideElement::Chart(chart) => {
            chart.chart.title.as_ref().is_some_and(|title| !title.trim().is_empty())
                || !chart.chart.series.is_empty()
                || !chart.chart.categories.is_empty()
        }
        SlideElement::Picture(_) | SlideElement::Media(_) => false,
    }
}

fn is_panel_shape(element: &SlideElement) -> bool {
    matches!(element, SlideElement::Shape(shape) if (shape.fill.is_some() || shape.stroke.is_some()) && shape.geometry != "line")
}

fn element_rect(element: &SlideElement) -> Option<Rect> {
    if let SlideElement::Shape(shape) = element {
        return shape_rect(shape);
    }
    let (x, y, width, height) = match element {
        SlideElement::Shape(_) => unreachable!(),
        SlideElement::Picture(value) => (value.x, value.y, value.width, value.height),
        SlideElement::Table(value) => (value.x, value.y, value.width, value.height),
        SlideElement::Chart(value) => (value.x, value.y, value.width, value.height),
        SlideElement::Media(value) => (value.x, value.y, value.width, value.height),
    };
    (width > 0 && height > 0).then_some(Rect {
        x: x as f64,
        y: y as f64,
        width: width as f64,
        height: height as f64,
    })
}

fn shape_rect(shape: &ShapeElement) -> Option<Rect> {
    (shape.width > 0 && shape.height > 0).then_some(Rect {
        x: shape.x as f64,
        y: shape.y as f64,
        width: shape.width as f64,
        height: shape.height as f64,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn block(order: usize, x: f64, y: f64, width: f64, height: f64) -> LayoutBlock {
        LayoutBlock {
            rect: Rect {
                x,
                y,
                width,
                height,
            },
            source_order: order,
            element_indices: vec![order],
            related: false,
            starts_new_region: false,
        }
    }

    #[test]
    fn repeated_cards_use_row_major_reading_order() {
        // Paint order alternates columns. Geometry should yield L1,L2,R1,R2.
        let blocks = vec![
            block(0, 0.0, 0.0, 40.0, 20.0),
            block(1, 60.0, 0.0, 40.0, 20.0),
            block(2, 0.0, 30.0, 40.0, 20.0),
            block(3, 60.0, 30.0, 40.0, 20.0),
        ];
        let ordered = spatial_order(blocks, 100.0, 50.0);
        let indices: Vec<usize> = ordered.iter().map(|block| block.source_order).collect();
        // A repeated 2x2 card layout is row-major, not column-major.
        assert_eq!(indices, vec![0, 1, 2, 3]);
    }

    #[test]
    fn tall_independent_columns_read_whole_left_then_whole_right() {
        let blocks = vec![
            block(0, 60.0, 0.0, 40.0, 80.0),
            block(1, 0.0, 0.0, 40.0, 20.0),
            block(2, 0.0, 30.0, 40.0, 50.0),
        ];
        let ordered = spatial_order(blocks, 100.0, 80.0);
        let indices: Vec<usize> = ordered.iter().map(|block| block.source_order).collect();
        assert_eq!(indices, vec![1, 2, 0]);
        assert!(ordered[2].starts_new_region);
    }

    #[test]
    fn asymmetric_columns_are_not_mistaken_for_a_repeated_grid() {
        let blocks = vec![
            block(0, 0.0, 5.0, 42.0, 70.0),
            block(1, 58.0, 0.0, 42.0, 20.0),
            block(2, 58.0, 30.0, 42.0, 20.0),
            block(3, 58.0, 60.0, 42.0, 20.0),
        ];
        let ordered = spatial_order(blocks, 100.0, 80.0);
        let indices: Vec<usize> = ordered.iter().map(|block| block.source_order).collect();
        assert_eq!(indices, vec![0, 1, 2, 3]);
        assert!(ordered[1].starts_new_region);
    }

    #[test]
    fn overlapping_backplates_form_one_connected_component() {
        let panels = vec![
            (
                2,
                Rect {
                    x: 10.0,
                    y: 0.0,
                    width: 40.0,
                    height: 40.0,
                },
            ),
            (
                3,
                Rect {
                    x: 0.0,
                    y: 30.0,
                    width: 40.0,
                    height: 40.0,
                },
            ),
            (
                4,
                Rect {
                    x: 30.0,
                    y: 30.0,
                    width: 40.0,
                    height: 40.0,
                },
            ),
            (
                9,
                Rect {
                    x: 80.0,
                    y: 0.0,
                    width: 20.0,
                    height: 20.0,
                },
            ),
        ];
        let mut components = connected_components(&panels, PANEL_COMPONENT_OVERLAP_RATIO);
        for component in &mut components {
            component.sort_unstable();
        }
        components.sort();
        assert_eq!(components, vec![vec![2, 3, 4], vec![9]]);
    }

    #[test]
    fn source_order_only_breaks_geometry_ties() {
        let blocks = vec![
            block(7, 0.0, 0.0, 10.0, 10.0),
            block(2, 0.0, 0.0, 10.0, 10.0),
        ];
        let ordered = spatial_order(blocks, 100.0, 100.0);
        assert_eq!(ordered[0].source_order, 2);
        assert_eq!(ordered[1].source_order, 7);
    }
}
