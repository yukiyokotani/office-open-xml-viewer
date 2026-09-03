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
}

/// Resource-governance ceiling for recursive whitespace partitioning. A slide
/// with more nested separators still gets deterministic geometry order, but a
/// hostile shape tree cannot turn Markdown projection into unbounded recursion.
const MAX_SPATIAL_PARTITION_DEPTH: usize = 64;

pub(crate) fn project_slide(slide: &Slide, slide_width: i64, slide_height: i64) -> SemanticSlide {
    let title = title_candidate(slide, slide_width, slide_height);
    let title_index = title.as_ref().map(|(index, _)| *index);
    let visible: Vec<usize> = slide
        .elements
        .iter()
        .enumerate()
        .filter_map(|(index, element)| {
            (Some(index) != title_index && element_has_content(element)).then_some(index)
        })
        .collect();

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

    // A common ungrouped authoring pattern is a filled/stroked rectangle
    // behind several text boxes. Associate content with the smallest authored
    // underlay that contains it. Full-slide backgrounds are deliberately
    // excluded: they describe the slide, not a local semantic block.
    let slide_area = (slide_width.max(1) as f64) * (slide_height.max(1) as f64);
    let mut memberships: HashMap<usize, Vec<usize>> = HashMap::new();
    for index in visible
        .iter()
        .copied()
        .filter(|index| !claimed.contains(index))
    {
        let Some(content_rect) = element_rect(&slide.elements[index]) else {
            continue;
        };
        let panel = slide
            .elements
            .iter()
            .enumerate()
            .filter(|(panel_index, element)| {
                *panel_index < index
                    && !element_has_content(element)
                    && is_panel_shape(element)
                    && element_rect(element).is_some_and(|rect| {
                        rect.area() < slide_area * PANEL_MAX_SLIDE_AREA_RATIO
                            && rect.contains(content_rect)
                    })
            })
            .filter_map(|(panel_index, element)| {
                element_rect(element).map(|rect| (panel_index, rect.area()))
            })
            .min_by(|a, b| a.1.partial_cmp(&b.1).unwrap_or(Ordering::Equal))
            .map(|(panel_index, _)| panel_index);
        if let Some(panel_index) = panel {
            memberships.entry(panel_index).or_default().push(index);
        }
    }
    for members in memberships
        .into_values()
        .filter(|members| members.len() >= 2)
    {
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
            // Repeated cards should read row-major. Outside that case, use the
            // stronger whitespace separator; source order is only a tie-breaker.
            let repeated_grid = blocks.len() >= 4 && h.left_len >= 2 && h.right_len >= 2;
            if repeated_grid || h.score >= v.score {
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
    result.extend(spatial_order_at_depth(
        second,
        slide_width,
        slide_height,
        depth + 1,
    ));
    result
}

#[derive(Clone, Copy)]
enum Axis {
    Horizontal,
    Vertical,
}

struct Split {
    first_orders: HashSet<usize>,
    left_len: usize,
    right_len: usize,
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
                first_orders: ordered[..cut]
                    .iter()
                    .map(|block| block.source_order)
                    .collect(),
                left_len: cut,
                right_len: ordered.len() - cut,
                score,
            });
        }
    }
    best
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
    let (x, y, width, height) = match element {
        SlideElement::Shape(value) => (value.x, value.y, value.width, value.height),
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
        let indices: Vec<usize> = ordered
            .into_iter()
            .map(|block| block.source_order)
            .collect();
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
        let indices: Vec<usize> = ordered
            .into_iter()
            .map(|block| block.source_order)
            .collect();
        assert_eq!(indices, vec![1, 2, 0]);
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
