//! Page layouts (ODF Part 3 §16.5, §17.2), master pages (§16.9), headers
//! and footers (§16.10-§16.12, §17.5), page columns (§17.12), notes and line
//! numbering configuration (§16.29, §16.30), mapped onto DOCX sections.
//!
//! ODF measures the page margins to the header's top and the footer's
//! bottom. LibreOffice's header height (`fo:min-height` or `svg:height`)
//! includes its spacing to the body (`fo:margin-bottom`), so the body starts
//! the height below the margin (measured: a 0.601 cm header with 0.501 cm
//! spacing starts the body 0.601 cm below the margin). DOCX measures the
//! header from the page edge (header distance) and the body from the page
//! edge (top margin), so the body margin is the page margin plus the
//! header's height. A header taller than its height pushes the body down
//! in both. LibreOffice keeps the spacing below it unless the region uses
//! dynamic spacing; DOCX never keeps it.

use super::*;
use docx_model::{
    ColSpec, ColumnsSpec, HeaderFooter, LineNumbering, NoteLayoutSettingsWire, NoteSeparatorMark,
    PageBorderEdge, PageBorders, PageNumType, SectionGeom, SectionPlacementWire, SectionProps,
};

/// The start of a DOCX section in the body.
#[derive(Clone)]
pub(super) struct SectionStart {
    pub master: String,
    /// How the section starts (ST_SectionMark): "nextPage", "continuous",
    /// "oddPage" or "evenPage".
    pub kind: &'static str,
    pub columns: Option<ColumnsSpec>,
    pub page_start: Option<i64>,
    /// Index of its marker block (informational).
    pub block: usize,
}

/// A page's text area.
#[derive(Clone, Copy)]
pub(super) struct Geometry {
    pub text_width: f64,
    pub text_height: f64,
    pub left: f64,
    pub right: f64,
}

/// A header or footer region's geometry (§17.5).
#[derive(Clone, Copy, Default, Debug)]
struct Region {
    /// Height of the region's content (min or fixed).
    height: f64,
    /// Spacing between the region and the body (`fo:margin-bottom` of a
    /// header, `fo:margin-top` of a footer), within `height`.
    spacing: f64,
    margin_left: f64,
    margin_right: f64,
}

/// Implementation bound on the columns of one page or section style.
const MAX_STYLE_COLUMNS: usize = 1024;

/// One page layout. Every field that owns data of the document's size
/// (fills, image link, column list) is shared, so a layout inherited from
/// the default page layout (§16.8) copies a fixed-size record, never the
/// default's data.
#[derive(Clone)]
struct Layout {
    section: SectionProps,
    columns: Option<Rc<ColumnsSpec>>,
    num_format: Option<&'static str>,
    usage: &'static str,
    header: Option<Region>,
    footer: Option<Region>,
    text_direction: Option<&'static str>,
    borders: Option<PageBorders>,
    grid: Option<(&'static str, f64, Option<f64>)>,
    footnote_separator: Option<NoteSeparatorMark>,
    /// The page's area fill (§20.175 fo:background-color, §20.120 draw:fill
    /// and companions on the page layout or its drawing page, or a
    /// `style:background-image`), held as a graphic style for the shared
    /// fill reader.
    background: Option<Rc<crate::style::Style>>,
    background_image: Option<Rc<(String, bool)>>,
}

struct Master<'a, 'i> {
    name: String,
    layout: Option<String>,
    next: Option<String>,
    node: roxmltree::Node<'a, 'i>,
}

pub(super) struct Pages<'a, 'i> {
    masters: Vec<Master<'a, 'i>>,
    by_name: HashMap<String, usize>,
    layouts: HashMap<String, Layout>,
    default_layout: Layout,
}

/// The header and footer stories of one master page as a DOCX section uses
/// them.
pub(super) struct MasterStories {
    pub headers: HeadersFooters,
    pub footers: HeadersFooters,
    pub title_page: bool,
    pub has_left: bool,
    /// Whether the page style shows a header / footer region.
    pub shown: [bool; 2],
    /// The construction bytes each story's tree was charged when it was
    /// built (headers, footers × default, first, even): every copy of a
    /// story is charged the same again before it is made.
    pub costs: [[usize; 3]; 2],
}

/// The fixed native size of one section start and of the copies its DOCX
/// section makes of it: the start itself (recorded and merged), the section
/// properties, geometry, placement and break, and the column list in each.
pub(super) fn section_bytes(start: &SectionStart) -> usize {
    let columns = start.columns.as_ref().map_or(0, |columns| {
        std::mem::size_of::<ColumnsSpec>() + columns.cols.len() * std::mem::size_of::<ColSpec>()
    });
    3 * (std::mem::size_of::<SectionStart>() + start.master.len())
        + 5 * columns
        + 2 * std::mem::size_of::<SectionProps>()
        + std::mem::size_of::<SectionPlacementWire>()
        + std::mem::size_of::<SectionGeom>()
        + std::mem::size_of::<BodyElement>()
        + 1024
}

fn default_section() -> SectionProps {
    // A4 with 2 cm margins (LibreOffice's default page).
    SectionProps {
        page_width: 595.3,
        page_height: 841.9,
        margin_top: 56.7,
        margin_right: 56.7,
        margin_bottom: 56.7,
        margin_left: 56.7,
        header_distance: 56.7,
        footer_distance: 56.7,
        ..SectionProps::default()
    }
}

impl<'a, 'i> Pages<'a, 'i> {
    pub fn read(styles_root: roxmltree::Node<'a, 'i>) -> Result<Self, String> {
        let mut default_layout = Layout {
            section: default_section(),
            columns: None,
            num_format: None,
            usage: "all",
            header: None,
            footer: None,
            text_direction: None,
            borders: None,
            grid: None,
            footnote_separator: None,
            background: None,
            background_image: None,
        };
        // §16.8: properties a page layout leaves unspecified come from the
        // default page layout.
        if let Some(default) = styles_root
            .children()
            .find(|n| xml::is(*n, NS_OFFICE, "styles"))
            .and_then(|s| {
                s.children()
                    .find(|n| xml::is(*n, NS_STYLE, "default-page-layout"))
            })
        {
            read_layout(default, &mut default_layout)?;
        }
        let mut layouts = HashMap::new();
        for automatic in styles_root
            .children()
            .filter(|n| xml::is(*n, NS_OFFICE, "automatic-styles"))
        {
            for node in automatic
                .children()
                .filter(|n| xml::is(*n, NS_STYLE, "page-layout"))
            {
                let Some(name) = node.attribute((NS_STYLE, "name")) else {
                    continue;
                };
                let mut layout = default_layout.clone();
                read_layout(node, &mut layout)?;
                layouts.insert(name.to_string(), layout);
            }
        }
        let mut masters = Vec::new();
        let mut by_name = HashMap::new();
        if let Some(styles) = styles_root
            .children()
            .find(|n| xml::is(*n, NS_OFFICE, "master-styles"))
        {
            for node in styles
                .children()
                .filter(|n| xml::is(*n, NS_STYLE, "master-page"))
            {
                let Some(name) = node.attribute((NS_STYLE, "name")) else {
                    continue;
                };
                by_name.entry(name.to_string()).or_insert(masters.len());
                masters.push(Master {
                    name: name.to_string(),
                    layout: node
                        .attribute((NS_STYLE, "page-layout-name"))
                        .map(str::to_string),
                    next: node
                        .attribute((NS_STYLE, "next-style-name"))
                        .filter(|next| *next != name)
                        .map(str::to_string),
                    node,
                });
            }
        }
        Ok(Self {
            masters,
            by_name,
            layouts,
            default_layout,
        })
    }

    pub fn has_master(&self, name: &str) -> bool {
        self.by_name.contains_key(name)
    }

    fn master(&self, name: &str) -> Option<&Master<'a, 'i>> {
        self.by_name.get(name).map(|index| &self.masters[*index])
    }

    fn layout(&self, master: &str) -> &Layout {
        self.master(master)
            .and_then(|m| m.layout.as_ref())
            .and_then(|name| self.layouts.get(name))
            .unwrap_or(&self.default_layout)
    }

    /// The master page whose layout the section's pages mostly use: a page
    /// style followed by another (`style:next-style-name`) lends its first
    /// page only.
    fn body_master<'s>(&'s self, master: &'s str) -> &'s str {
        self.master(master)
            .and_then(|m| m.next.as_deref())
            .filter(|next| self.has_master(next))
            .unwrap_or(master)
    }

    /// The initial page style: the first paragraph's, else LibreOffice's
    /// "Standard", else the first master page.
    pub fn initial_master(&self, body: roxmltree::Node<'_, '_>) -> &str {
        let _ = body;
        if self.has_master("Standard") {
            return "Standard";
        }
        self.masters.first().map_or("Standard", |m| m.name.as_str())
    }

    pub fn columns(&self, master: &str) -> Option<ColumnsSpec> {
        self.layout(self.body_master(master))
            .columns
            .as_deref()
            .cloned()
    }

    /// How a section with this page style starts: on a right (odd) page for
    /// a right-only style, on a left one for a left-only style (§19.493
    /// style:page-usage).
    pub fn start_kind(&self, master: &str) -> &'static str {
        match self.layout(master).usage {
            "right" => "oddPage",
            "left" => "evenPage",
            _ => "nextPage",
        }
    }

    pub fn geometry(&self, master: &str) -> Geometry {
        let section = &self.layout(self.body_master(master)).section;
        Geometry {
            text_width: (section.page_width - section.margin_left - section.margin_right).max(1.0),
            text_height: (section.page_height - section.margin_top - section.margin_bottom)
                .max(1.0),
            left: section.margin_left,
            right: section.margin_right,
        }
    }

    fn mirrored(&self) -> bool {
        self.layouts
            .values()
            .any(|layout| layout.usage == "mirrored")
    }

    /// Horizontal writing without a layout grid on the page style's body
    /// pages: the scope in which LibreOffice's fixed-line placement was
    /// measured.
    pub fn plain_horizontal(&self, master: &str) -> bool {
        let layout = self.layout(self.body_master(master));
        layout.text_direction.is_none() && layout.grid.is_none()
    }

    /// Horizontal writing on the page style's body pages, grid or not.
    pub fn horizontal(&self, master: &str) -> bool {
        self.layout(self.body_master(master))
            .text_direction
            .is_none()
    }

    pub fn footnote_separator(&self, master: &str) -> Option<NoteSeparatorMark> {
        self.layout(self.body_master(master)).footnote_separator
    }
}

/// One page layout's properties.
fn read_layout(node: roxmltree::Node<'_, '_>, layout: &mut Layout) -> Result<(), String> {
    // §19.493 style:page-usage; other values read as "all".
    if let Some(usage) = node.attribute((NS_STYLE, "page-usage")) {
        layout.usage = match usage {
            "left" => "left",
            "right" => "right",
            "mirrored" => "mirrored",
            _ => "all",
        };
    }
    for child in node.children().filter(|n| n.is_element()) {
        match (child.tag_name().namespace(), child.tag_name().name()) {
            (Some(NS_STYLE), "page-layout-properties") => read_page_properties(child, layout)?,
            (Some(NS_STYLE), "header-style") => {
                layout.header = region(child, "margin-bottom");
            }
            (Some(NS_STYLE), "footer-style") => {
                layout.footer = region(child, "margin-top");
            }
            _ => {}
        }
    }
    Ok(())
}

fn region(node: roxmltree::Node<'_, '_>, spacing: &str) -> Option<Region> {
    let properties = node
        .children()
        .find(|n| xml::is(*n, NS_STYLE, "header-footer-properties"))?;
    let get = |ns: &str, name: &str| {
        properties
            .attribute((ns, name))
            .and_then(|v| length_pt(v).ok())
    };
    Some(Region {
        height: get(NS_SVG, "height")
            .or_else(|| get(NS_FO, "min-height"))
            .unwrap_or(0.0)
            .max(0.0),
        // With dynamic spacing (§20.354, the Word-compatible mode) a grown
        // region eats into its spacing, which DOCX already does; otherwise
        // LibreOffice keeps the spacing below the grown content.
        spacing: if properties.attribute((NS_STYLE, "dynamic-spacing")) == Some("true") {
            0.0
        } else {
            get(NS_FO, spacing).unwrap_or(0.0).max(0.0)
        },
        margin_left: get(NS_FO, "margin-left").unwrap_or(0.0),
        margin_right: get(NS_FO, "margin-right").unwrap_or(0.0),
    })
}

fn read_page_properties(node: roxmltree::Node<'_, '_>, layout: &mut Layout) -> Result<(), String> {
    read_background(node, layout);
    let section = &mut layout.section;
    let get = |ns: &str, name: &str| node.attribute((ns, name)).and_then(|v| length_pt(v).ok());
    if let Some(width) = get(NS_FO, "page-width").filter(|v| *v > 0.0) {
        section.page_width = width;
    }
    if let Some(height) = get(NS_FO, "page-height").filter(|v| *v > 0.0) {
        section.page_height = height;
    }
    let all = get(NS_FO, "margin");
    for (name, slot) in [
        ("margin-top", &mut section.margin_top),
        ("margin-bottom", &mut section.margin_bottom),
        ("margin-left", &mut section.margin_left),
        ("margin-right", &mut section.margin_right),
    ] {
        if let Some(value) = get(NS_FO, name).or(all) {
            *slot = value.max(0.0);
        }
    }
    // A gutter (LibreOffice `loext:margin-gutter`) widens the left margin.
    if let Some(gutter) = get(NS_LOEXT, "margin-gutter").filter(|g| *g > 0.0) {
        section.margin_left += gutter;
    }
    if let Some(format) = node.attribute((NS_STYLE, "num-format")) {
        layout.num_format = page_number_format(format);
    }
    layout.text_direction = match node.attribute((NS_STYLE, "writing-mode")) {
        Some("tb-rl" | "tb") => Some("tbRl"),
        _ => layout.text_direction.take(),
    };
    // §20.183 page border and §20.217 padding: the text sits inside both.
    let border = |side: &str| {
        node.attribute((NS_FO, &*format!("border-{side}")))
            .or_else(|| node.attribute((NS_FO, "border")))
            .and_then(|value| {
                parse_border(
                    value,
                    node.attribute((NS_STYLE, &*format!("border-line-width-{side}")))
                        .or_else(|| node.attribute((NS_STYLE, "border-line-width"))),
                )
            })
    };
    let padding = |side: &str| {
        get(NS_FO, &format!("padding-{side}"))
            .or_else(|| get(NS_FO, "padding"))
            .unwrap_or(0.0)
            .max(0.0)
    };
    let mut edges = Vec::new();
    for side in ["top", "left", "bottom", "right"] {
        let edge = border(side).map(|spec| PageBorderEdge {
            style: spec.style.to_string(),
            color: spec.color,
            width: spec.width,
            space: padding(side),
        });
        let inset = edge.as_ref().map_or(0.0, |e| e.width) + padding(side);
        match side {
            "top" => section.margin_top += inset,
            "left" => section.margin_left += inset,
            "bottom" => section.margin_bottom += inset,
            _ => section.margin_right += inset,
        }
        edges.push(edge);
    }
    if edges.iter().any(Option::is_some) {
        let mut edges = edges.into_iter();
        layout.borders = Some(PageBorders {
            offset_from: "text".into(),
            display: "allPages".into(),
            z_order: "front".into(),
            top: edges.next().flatten(),
            left: edges.next().flatten(),
            bottom: edges.next().flatten(),
            right: edges.next().flatten(),
        });
    }
    // §20.309 layout grid.
    match node.attribute((NS_STYLE, "layout-grid-mode")) {
        Some(mode @ ("line" | "both")) => {
            let base = get(NS_STYLE, "layout-grid-base-height").unwrap_or(0.0);
            // §20.307: each grid line holds the base text and its ruby.
            let ruby = get(NS_STYLE, "layout-grid-ruby-height").unwrap_or(0.0);
            let char_width = get(NS_STYLE, "layout-grid-base-width");
            // §20.313 style:layout-grid-snap-to: text snaps to the
            // character grid only when true (the default).
            let chars = mode == "both"
                && node.attribute((NS_STYLE, "layout-grid-snap-to")) != Some("false");
            if base + ruby > 0.0 {
                layout.grid = Some((
                    if chars { "linesAndChars" } else { "lines" },
                    base + ruby,
                    if chars { char_width } else { None },
                ));
            }
        }
        Some("none") => layout.grid = None,
        _ => {}
    }
    for child in node.children().filter(|n| n.is_element()) {
        match (child.tag_name().namespace(), child.tag_name().name()) {
            (Some(NS_STYLE), "columns") => {
                let width =
                    (section.page_width - section.margin_left - section.margin_right).max(1.0);
                let captured = crate::style::Element {
                    namespace: NS_STYLE.into(),
                    name: "columns".into(),
                    attributes: crate::style::attributes(child),
                    children: child
                        .children()
                        .filter(|n| n.is_element())
                        .map(|c| crate::style::Element {
                            namespace: c.tag_name().namespace().unwrap_or("").into(),
                            name: c.tag_name().name().into(),
                            attributes: crate::style::attributes(c),
                            children: Vec::new(),
                        })
                        .collect(),
                };
                layout.columns = columns(&captured, width)?.map(Rc::new);
            }
            (Some(NS_STYLE), "footnote-sep") => {
                let rel = child
                    .attribute((NS_STYLE, "rel-width"))
                    .and_then(|v| percent(v).ok())
                    .unwrap_or(0.25);
                let width = child
                    .attribute((NS_STYLE, "width"))
                    .and_then(|v| length_pt(v).ok())
                    .unwrap_or(0.5);
                layout.footnote_separator = Some(
                    if child.attribute((NS_STYLE, "line-style")) == Some("none") || width <= 0.0 {
                        NoteSeparatorMark::None
                    } else if rel >= 0.99 {
                        NoteSeparatorMark::Full
                    } else {
                        NoteSeparatorMark::Short
                    },
                );
            }
            _ => {}
        }
    }
    Ok(())
}

/// A page's area fill: fo:background-color, the draw:fill attributes and a
/// background image, from page layout or drawing page properties.
fn read_background(node: roxmltree::Node<'_, '_>, layout: &mut Layout) {
    let attributes: HashMap<(String, String), String> = crate::style::attributes(node)
        .into_iter()
        .filter(|((ns, name), _)| {
            (ns == NS_FO && name == "background-color")
                || (ns == NS_DRAW && (name.starts_with("fill") || name == "opacity"))
                || (ns == NS_STYLE && name == "repeat")
        })
        .collect();
    let visible = attributes
        .get(&(NS_FO.to_string(), "background-color".to_string()))
        .is_some_and(|v| v != "transparent")
        || attributes
            .get(&(NS_DRAW.to_string(), "fill".to_string()))
            .is_some_and(|v| v != "none");
    if visible {
        layout.background = Some(Rc::new(crate::style::Style {
            name: String::new(),
            family: "graphic".into(),
            parent: None,
            automatic: true,
            attributes: HashMap::new(),
            groups: HashMap::from([(
                "graphic-properties".to_string(),
                crate::style::Properties {
                    attributes,
                    ..crate::style::Properties::default()
                },
            )]),
        }));
    }
    if let Some(image) = node
        .children()
        .find(|n| xml::is(*n, NS_STYLE, "background-image"))
        .filter(|n| n.attribute((NS_XLINK, "href")).is_some())
    {
        layout.background_image = Some(Rc::new((
            image
                .attribute((NS_XLINK, "href"))
                .unwrap_or("")
                .to_string(),
            image.attribute((NS_STYLE, "repeat")) != Some("stretch"),
        )));
    }
}

/// `style:num-format` (§19.500) as an ST_NumberFormat value.
fn page_number_format(format: &str) -> Option<&'static str> {
    match format {
        "1" => Some("decimal"),
        "i" => Some("lowerRoman"),
        "I" => Some("upperRoman"),
        "a" => Some("lowerLetter"),
        "A" => Some("upperLetter"),
        _ => None,
    }
}

/// `style:columns` (§17.12): equal columns from `fo:column-count` and
/// `fo:column-gap`, or the `style:column` children's relative widths and
/// indents within `width`; a separator line (`style:column-sep`) maps to the
/// DOCX separator.
/// More than `MAX_STYLE_COLUMNS` columns is beyond the reader's limits.
pub(super) fn columns(
    element: &crate::style::Element,
    width: f64,
) -> Result<Option<ColumnsSpec>, String> {
    let attribute = |ns: &str, name: &str| {
        element
            .attributes
            .get(&(ns.to_string(), name.to_string()))
            .map(String::as_str)
    };
    let count = attribute(NS_FO, "column-count")
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(1);
    if count <= 1 {
        return Ok(None);
    }
    if count > MAX_STYLE_COLUMNS {
        return Err("OpenDocument columns exceed the column limit".into());
    }
    let gap = attribute(NS_FO, "column-gap")
        .and_then(|v| length_pt(v).ok())
        .unwrap_or(0.0);
    let sep = element.children.iter().any(|child| {
        child.name == "column-sep"
            && child
                .attributes
                .get(&(NS_STYLE.to_string(), "style".to_string()))
                .is_none_or(|style| style != "none")
    });
    let columns: Vec<&crate::style::Element> = element
        .children
        .iter()
        .filter(|child| child.name == "column")
        .collect();
    if columns.len() == count {
        fn get<'e>(column: &'e crate::style::Element, ns: &str, name: &str) -> Option<&'e str> {
            column
                .attributes
                .get(&(ns.to_string(), name.to_string()))
                .map(String::as_str)
        }
        let weights: Vec<f64> = columns
            .iter()
            .map(|c| {
                get(c, NS_STYLE, "rel-width")
                    .map(|v| v.trim_end_matches('*'))
                    .and_then(|v| v.parse::<f64>().ok())
                    .unwrap_or(1.0)
            })
            .collect();
        let total: f64 = weights.iter().sum();
        let indents: Vec<(f64, f64)> = columns
            .iter()
            .map(|c| {
                (
                    get(c, NS_FO, "start-indent")
                        .and_then(|v| length_pt(v).ok())
                        .unwrap_or(0.0),
                    get(c, NS_FO, "end-indent")
                        .and_then(|v| length_pt(v).ok())
                        .unwrap_or(0.0),
                )
            })
            .collect();
        if total > 0.0 {
            let mut cols = Vec::with_capacity(count);
            for i in 0..count {
                let outer = weights[i] / total * width;
                let text = (outer - indents[i].0 - indents[i].1).max(1.0);
                let space = if i + 1 < count {
                    indents[i].1 + indents[i + 1].0
                } else {
                    0.0
                };
                cols.push(ColSpec {
                    width_pt: text,
                    space_pt: space,
                });
            }
            let equal = cols
                .windows(2)
                .all(|pair| (pair[0].width_pt - pair[1].width_pt).abs() < 0.05)
                && cols[..count - 1]
                    .windows(2)
                    .all(|pair| (pair[0].space_pt - pair[1].space_pt).abs() < 0.05);
            let space = cols.first().map_or(gap, |c| c.space_pt);
            return Ok(Some(ColumnsSpec {
                count,
                space_pt: space,
                equal_width: equal,
                sep,
                cols,
            }));
        }
    }
    let column_width = ((width - gap * (count - 1) as f64) / count as f64).max(1.0);
    Ok(Some(ColumnsSpec {
        count,
        space_pt: gap,
        equal_width: true,
        sep,
        cols: (0..count)
            .map(|i| ColSpec {
                width_pt: column_width,
                space_pt: if i + 1 < count { gap } else { 0.0 },
            })
            .collect(),
    }))
}

/// `text:linenumbering-configuration` (§16.30) as DOCX line numbering.
pub(super) fn line_numbering(styles: Option<roxmltree::Node<'_, '_>>) -> Option<LineNumbering> {
    let node = styles?
        .children()
        .find(|n| xml::is(*n, NS_TEXT, "linenumbering-configuration"))?;
    if node.attribute((NS_TEXT, "number-lines")) == Some("false") {
        return None;
    }
    Some(LineNumbering {
        count_by: node
            .attribute((NS_TEXT, "increment"))
            .and_then(|v| v.parse::<i64>().ok())
            .filter(|v| *v > 0)
            .unwrap_or(1),
        start: 1,
        distance: node
            .attribute((NS_TEXT, "offset"))
            .and_then(|v| length_pt(v).ok()),
        restart: if node.attribute((NS_TEXT, "restart-on-page")) == Some("true") {
            "newPage".into()
        } else {
            "continuous".into()
        },
    })
}

/// `text:notes-configuration` (§16.29).
#[derive(Default)]
pub(super) struct NotesConfiguration {
    footnote: Option<NoteConfig>,
    endnote: Option<NoteConfig>,
}

#[derive(Default)]
struct NoteConfig {
    format: Option<&'static str>,
    start: Option<i64>,
    citation: Option<String>,
    body_citation: Option<String>,
    at_document_end: bool,
}

impl NotesConfiguration {
    pub fn read(styles: Option<roxmltree::Node<'_, '_>>) -> Self {
        let mut out = Self::default();
        let Some(styles) = styles else {
            return out;
        };
        for node in styles
            .children()
            .filter(|n| xml::is(*n, NS_TEXT, "notes-configuration"))
        {
            let config = NoteConfig {
                format: node
                    .attribute((NS_STYLE, "num-format"))
                    .and_then(|f| list::note_number_format(f)),
                // LibreOffice writes the offset from 1 (0 = start at 1).
                start: node
                    .attribute((NS_TEXT, "start-value"))
                    .and_then(|v| v.parse::<i64>().ok())
                    .map(|v| v + 1),
                citation: node
                    .attribute((NS_TEXT, "citation-style-name"))
                    .map(str::to_string),
                body_citation: node
                    .attribute((NS_TEXT, "citation-body-style-name"))
                    .map(str::to_string),
                at_document_end: node.attribute((NS_TEXT, "footnotes-position"))
                    == Some("document"),
            };
            match node.attribute((NS_TEXT, "note-class")) {
                Some("endnote") => out.endnote = Some(config),
                _ => out.footnote = Some(config),
            }
        }
        out
    }

    fn config(&self, kind: &str) -> Option<&NoteConfig> {
        if kind == "endnote" {
            self.endnote.as_ref()
        } else {
            self.footnote.as_ref()
        }
    }

    /// The text style of a note's number inside the note (§19.774); without
    /// one, LibreOffice's default "Footnote Symbol" / "Endnote Symbol".
    pub fn citation_style(&self, kind: &str) -> &str {
        self.config(kind)
            .and_then(|c| c.citation.as_deref())
            .unwrap_or(if kind == "endnote" {
                "Endnote_20_Symbol"
            } else {
                "Footnote_20_Symbol"
            })
    }

    /// The text style of a note's citation in the text flow (§19.773);
    /// without one, LibreOffice's default "Footnote anchor" / "Endnote
    /// anchor".
    pub fn body_citation_style(&self, kind: &str) -> &str {
        self.config(kind)
            .and_then(|c| c.body_citation.as_deref())
            .unwrap_or(if kind == "endnote" {
                "Endnote_20_anchor"
            } else {
                "Footnote_20_anchor"
            })
    }

    pub fn wire(&self, pages: &Pages<'_, '_>, master: &str) -> Option<NoteLayoutSettingsWire> {
        let footnote = self.footnote.as_ref();
        let endnote = self.endnote.as_ref();
        Some(NoteLayoutSettingsWire {
            footnote_position: footnote
                .filter(|c| c.at_document_end)
                .map(|_| "docEnd".to_string()),
            endnote_position: Some("docEnd".to_string()),
            footnote_number_format: footnote.and_then(|c| c.format).map(str::to_string),
            footnote_number_start: footnote.and_then(|c| c.start),
            endnote_number_format: Some(
                endnote
                    .and_then(|c| c.format)
                    .unwrap_or("lowerRoman")
                    .to_string(),
            ),
            endnote_number_start: endnote.and_then(|c| c.start),
            footnote_separator: pages.footnote_separator(master),
            ..NoteLayoutSettingsWire::default()
        })
    }
}

impl<'a> Reader<'a> {
    /// The header and footer stories of a master page as a section uses
    /// them, built once per master page.
    fn master_stories(&mut self, master: &str) -> Result<Rc<MasterStories>, String> {
        if let Some(stories) = self.headers.get(master) {
            return Ok(stories.clone());
        }
        let pages = self.pages;
        let body_master = pages.body_master(master).to_string();
        let mut stories = self.region_stories(&body_master)?;
        if body_master != master {
            // A first-page style followed by the body style: its own
            // regions are the section's first-page ones.
            // Moved out of the first style's own stories, not copied.
            let mut first = self.region_stories(master)?;
            stories.headers.first = first
                .headers
                .default
                .take()
                .or(Some(HeaderFooter::default()));
            stories.footers.first = first
                .footers
                .default
                .take()
                .or(Some(HeaderFooter::default()));
            stories.costs[0][1] = first.costs[0][0];
            stories.costs[1][1] = first.costs[1][0];
            stories.title_page = true;
        }
        let stories = Rc::new(stories);
        self.headers.insert(master.to_string(), stories.clone());
        Ok(stories)
    }

    fn region_stories(&mut self, master: &str) -> Result<MasterStories, String> {
        let pages = self.pages;
        let Some(master_page) = pages.master(master) else {
            return Ok(MasterStories {
                headers: HeadersFooters::default(),
                footers: HeadersFooters::default(),
                title_page: false,
                has_left: false,
                shown: [false; 2],
                costs: [[0; 3]; 2],
            });
        };
        let node = master_page.node;
        let layout = pages.layout(master);
        let geometry = pages.geometry(master);
        let mut out = MasterStories {
            headers: HeadersFooters::default(),
            footers: HeadersFooters::default(),
            title_page: false,
            has_left: false,
            shown: [false; 2],
            costs: [[0; 3]; 2],
        };
        for (kind, region) in [("header", layout.header), ("footer", layout.footer)] {
            let shown = |name: &str| {
                node.children().find(|n| {
                    xml::is(*n, NS_STYLE, name)
                        && n.attribute((NS_STYLE, "display")) != Some("false")
                })
            };
            let Some(default) = shown(kind) else {
                // No region: an empty story so that the section does not
                // inherit the previous section's (DOCX inheritance).
                let empty = HeadersFooters {
                    default: Some(HeaderFooter::default()),
                    first: None,
                    even: None,
                };
                if kind == "header" {
                    out.headers = empty;
                } else {
                    out.footers = empty;
                }
                continue;
            };
            let which = usize::from(kind == "footer");
            out.shown[which] = true;
            let region = region.unwrap_or_default();
            let width = geometry.text_width - region.margin_left - region.margin_right;
            // A story and the construction bytes its tree was charged.
            let build = |reader: &mut Self,
                         node: roxmltree::Node<'a, 'a>|
             -> Result<(HeaderFooter, usize), String> {
                let before = reader.copyable_bytes();
                let saved_scope = std::mem::replace(&mut reader.scope, Scope::Styles);
                let story = reader.contained_story(node, width);
                reader.scope = saved_scope;
                let mut body = story?;
                for element in &mut body {
                    if let BodyElement::Paragraph(p) = element {
                        p.indent_left += region.margin_left;
                        p.indent_right += region.margin_right;
                    }
                }
                // LibreOffice keeps the region's spacing between a header
                // taller than its height and the body (and between the body
                // and a footer): the spacing joins the region's content, so a
                // DOCX region that grows past the body margin pushes the body
                // by its content plus the spacing.
                let edge = if kind == "header" {
                    body.iter_mut().rev().find_map(|e| match e {
                        BodyElement::Paragraph(p) => Some(p),
                        _ => None,
                    })
                } else {
                    body.iter_mut().find_map(|e| match e {
                        BodyElement::Paragraph(p) => Some(p),
                        _ => None,
                    })
                };
                if let Some(paragraph) = edge {
                    if kind == "header" {
                        paragraph.space_after += region.spacing;
                    } else {
                        paragraph.space_before += region.spacing;
                    }
                }
                Ok((
                    HeaderFooter { body },
                    reader.copyable_bytes().saturating_sub(before),
                ))
            };
            // Lifetimes: master nodes live as long as the styles document.
            let default_node: roxmltree::Node<'a, 'a> = unsafe_node(default);
            let (story, cost) = build(self, default_node)?;
            out.costs[which][0] = cost;
            let left = shown(&format!("{kind}-left")).map(unsafe_node);
            let first = shown(&format!("{kind}-first")).map(unsafe_node);
            let even = match left {
                Some(left) => {
                    out.has_left = true;
                    let (story, cost) = build(self, left)?;
                    out.costs[which][2] = cost;
                    Some(story)
                }
                None => None,
            };
            let first_story = match first {
                Some(first) => {
                    out.title_page = true;
                    let (story, cost) = build(self, first)?;
                    out.costs[which][1] = cost;
                    Some(story)
                }
                None => None,
            };
            let set = HeadersFooters {
                default: Some(story),
                first: first_story,
                even,
            };
            if kind == "header" {
                out.headers = set;
            } else {
                out.footers = set;
            }
        }
        // The page background, drawn behind the text on every page by an
        // anchored, page-sized shape in each header story.
        let before = self.copyable_bytes();
        if let Some(background) = self.page_background(layout)? {
            // Each header story owns a copy of the background shape and its
            // paragraph: every copy is charged before it is made.
            let copy = self.copyable_bytes().saturating_sub(before)
                + std::mem::size_of::<DocParagraph>()
                + std::mem::size_of::<BodyElement>();
            let paragraph =
                || BodyElement::Paragraph(Box::new(anchor_paragraph(vec![background.clone()])));
            self.charge_construction(copy)?;
            out.costs[0][0] += copy;
            out.headers
                .default
                .get_or_insert_with(HeaderFooter::default)
                .body
                .insert(0, paragraph());
            for (slot, story) in [(1, &mut out.headers.first), (2, &mut out.headers.even)] {
                if let Some(story) = story {
                    self.charge_construction(copy)?;
                    out.costs[0][slot] += copy;
                    story.body.insert(0, paragraph());
                }
            }
        }
        // A first page with only one region's first variant uses the other
        // region's ordinary story.
        if out.title_page {
            for (which, set) in [(0, &mut out.headers), (1, &mut out.footers)] {
                if set.first.is_none() {
                    self.charge_construction(out.costs[which][0])?;
                    out.costs[which][1] = out.costs[which][0];
                    set.first = set.default.clone();
                }
            }
        }
        Ok(out)
    }

    /// The DOCX section properties of a section start.
    fn section_parts(
        &mut self,
        start: &SectionStart,
        line_numbering: &Option<LineNumbering>,
        ordinal: usize,
        even_and_odd: bool,
    ) -> Result<
        (
            SectionProps,
            HeadersFooters,
            HeadersFooters,
            bool,
            SectionPlacementWire,
        ),
        String,
    > {
        let pages = self.pages;
        let body_master = pages.body_master(&start.master).to_string();
        let layout = pages.layout(&body_master).clone();
        let stories = self.master_stories(&start.master)?;
        let mut section = layout.section.clone();
        // Header and footer regions: the body starts below the header's
        // height (see the module documentation).
        section.header_distance = section.margin_top;
        section.footer_distance = section.margin_bottom;
        if let Some(header) = layout.header.filter(|_| stories.shown[0]) {
            section.margin_top += header.height;
        }
        // A header story that only carries the page background sits at the
        // page's top edge, where its empty line cannot reach the body.
        if !stories.shown[0] {
            section.header_distance = 0.0;
        }
        if let Some(footer) = layout.footer.filter(|_| stories.shown[1]) {
            section.margin_bottom += footer.height;
        }
        section.title_page = stories.title_page;
        section.section_start = Some(start.kind.to_string());
        section.columns = start.columns.clone();
        section.text_direction = layout.text_direction.map(str::to_string);
        section.page_num_type =
            (layout.num_format.is_some() || start.page_start.is_some()).then(|| PageNumType {
                start: start.page_start,
                fmt: layout.num_format.map(str::to_string),
            });
        section.page_borders = layout.borders.clone();
        section.line_numbering = line_numbering.clone();
        if let Some((kind, pitch, chars)) = layout.grid {
            section.doc_grid_type = Some(kind.into());
            section.doc_grid_line_pitch = Some(pitch);
            section.doc_grid_char_space = chars.map(|width| {
                let base = self.default_font_size();
                ((width - base) * 4096.0).round()
            });
        }
        // Every section owns its copies of the master's stories (DOCX
        // sections resolve their header/footer references independently):
        // each copy is charged its story's construction cost first.
        let set_cost = |costs: &[usize; 3]| costs.iter().fold(0usize, |a, c| a.saturating_add(*c));
        self.charge_construction(
            set_cost(&stories.costs[0]).saturating_add(set_cost(&stories.costs[1])),
        )?;
        let mut headers = stories.headers.clone();
        let mut footers = stories.footers.clone();
        if even_and_odd {
            for (which, set) in [(0, &mut headers), (1, &mut footers)] {
                if set.even.is_none() {
                    self.charge_construction(stories.costs[which][0])?;
                    set.even = set.default.clone();
                }
            }
        }
        let placement = SectionPlacementWire {
            section_id: format!("section:{ordinal}"),
            section_bidi: false,
            v_align: None,
            line_numbering: section.line_numbering.clone(),
            doc_grid_type: section.doc_grid_type.clone(),
            doc_grid_line_pitch: section.doc_grid_line_pitch,
            doc_grid_char_space: section.doc_grid_char_space,
            gutter_pt: None,
            rtl_gutter: None,
            section_flow: None,
            page_borders_authored: Some(section.page_borders.is_some()),
            page_borders: section.page_borders.clone(),
            page_geometry: None,
        };
        Ok((section, headers, footers, stories.title_page, placement))
    }

    /// The page background as a behind-text shape covering the page.
    fn page_background(&mut self, layout: &'a Layout) -> Result<Option<DocRun>, String> {
        let mut fill = match &layout.background {
            Some(style) => self.background_fill(style)?,
            None => None,
        };
        if fill.is_none() {
            if let Some((href, tile)) = layout.background_image.as_deref() {
                if let Some((path, mime)) = self.package_image(href)? {
                    fill = Some(docx_model::ShapeFill::Image {
                        image_path: path,
                        mime_type: mime.into(),
                        svg_image_path: None,
                        src_rect: None,
                        fill_rect: None,
                        tile: tile.then(|| {
                            Box::new(ooxml_common::fill::TileInfo {
                                tx: None,
                                ty: None,
                                sx: None,
                                sy: None,
                                flip: None,
                                algn: None,
                            })
                        }),
                        alpha: None,
                        duotone: None,
                        blip_effects: Vec::new(),
                        vml: None,
                    });
                }
            }
        }
        let Some(fill) = fill else {
            return Ok(None);
        };
        Ok(Some(self.page_background_shape(
            fill,
            layout.section.page_width,
            layout.section.page_height,
        )))
    }

    fn default_font_size(&self) -> f64 {
        self.sheet
            .default_style(Family::Paragraph)
            .and_then(|style| text::paragraph_font_size(Cascade::new(vec![style], TEXT)))
            .unwrap_or(12.0)
    }

    /// The final document: body elements with the section breaks between
    /// sections, the last section in `Document::section`.
    pub(super) fn assemble(
        &mut self,
        body: Vec<BodyElement>,
        marks: Vec<(usize, usize)>,
        line_numbering: Option<LineNumbering>,
    ) -> Result<Document, String> {
        // Section starts by body position; starts at one position merge into
        // the last (a page start among them wins).
        let mut starts: Vec<(usize, SectionStart)> = vec![(0, self.sections[0].clone())];
        for (position, index) in marks {
            let mut start = self.sections[index].clone();
            if let Some(last) = starts.last_mut() {
                if last.0 == position {
                    if start.kind == "continuous" && last.1.kind != "continuous" {
                        start.kind = last.1.kind;
                        if start.page_start.is_none() {
                            start.page_start = last.1.page_start;
                        }
                    }
                    last.1 = start;
                    continue;
                }
            }
            starts.push((position, start));
        }
        // Even/odd headers are document-wide in DOCX.
        let mut even_and_odd = false;
        for (_, start) in &starts {
            let master = start.master.clone();
            if self.master_stories(&master)?.has_left {
                even_and_odd = true;
            }
        }
        let embedded_fonts = self.embedded_fonts()?;
        let mut out = Vec::with_capacity(body.len() + starts.len());
        let mut body = body.into_iter().enumerate().peekable();
        let mut ordinal = 0;
        let mut section_index = 0;
        let count = starts.len();
        let mut final_parts = None;
        while section_index < count {
            let end = starts
                .get(section_index + 1)
                .map_or(usize::MAX, |(position, _)| *position);
            while let Some((index, _)) = body.peek() {
                if *index >= end {
                    break;
                }
                let (_, element) = body.next().expect("peeked");
                out.push(element);
            }
            let start = starts[section_index].1.clone();
            let parts = self.section_parts(&start, &line_numbering, ordinal, even_and_odd)?;
            if section_index + 1 < count {
                let (section, headers, footers, title_page, placement) = parts;
                out.push(BodyElement::SectionBreak {
                    kind: start.kind.to_string(),
                    columns: section.columns.clone(),
                    headers: Box::new(headers),
                    footers: Box::new(footers),
                    title_page,
                    geom: Some(Box::new(SectionGeom {
                        page_width: section.page_width,
                        page_height: section.page_height,
                        margin_top: section.margin_top,
                        margin_right: section.margin_right,
                        margin_bottom: section.margin_bottom,
                        margin_left: section.margin_left,
                        header_distance: section.header_distance,
                        footer_distance: section.footer_distance,
                    })),
                    page_num_type: section.page_num_type.clone(),
                    text_direction: section.text_direction.clone(),
                    section_placement: Box::new(placement),
                });
                ordinal += 1;
            } else {
                final_parts = Some(parts);
            }
            section_index += 1;
        }
        out.extend(body.map(|(_, element)| element));
        let (mut section, headers, footers, _, placement) =
            final_parts.expect("at least one section");
        section.even_and_odd_headers = even_and_odd;
        section.section_placement = Some(Box::new(placement));
        Ok(Document {
            section,
            body: out,
            headers,
            footers,
            embedded_fonts,
            page_layout_settings: self.pages.mirrored().then(|| {
                docx_model::PageLayoutSettingsWire {
                    mirror_margins: Some(true),
                    ..docx_model::PageLayoutSettingsWire::default()
                }
            }),
            ..Document::default()
        })
    }
}

/// Master-page nodes come from the styles document, which outlives the
/// reader like the content document; both are borrowed for `'a`.
fn unsafe_node<'a>(node: roxmltree::Node<'a, 'a>) -> roxmltree::Node<'a, 'a> {
    node
}
