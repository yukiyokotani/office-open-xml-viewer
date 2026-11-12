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
    /// The page's area fill properties (§20.175 fo:background-color,
    /// §20.120 draw:fill and companions), held as a graphic style for the
    /// shared fill reader: the default page layout's (§16.8) with the
    /// layout's own specified over them, property by property. An explicit
    /// `transparent` or `none` is a specified value like any other.
    background: Option<Rc<crate::style::Style>>,
    /// `style:background-image` (§17.3): its link and whether it tiles. A
    /// layout's own element replaces the default's, also when it links
    /// nothing.
    background_image: Option<Rc<(String, bool)>>,
}

/// A master page with its page layout and next style resolved once (as
/// indices), so later lookups never rehash the linked names.
struct Master<'a, 'i> {
    name: String,
    layout: Option<usize>,
    next: Option<usize>,
    node: roxmltree::Node<'a, 'i>,
}

pub(super) struct Pages<'a, 'i> {
    masters: Vec<Master<'a, 'i>>,
    by_name: HashMap<String, usize>,
    layouts: Vec<Layout>,
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
        let mut layout_index: HashMap<String, usize> = HashMap::new();
        let mut layouts: Vec<Layout> = Vec::new();
        for automatic in styles_root
            .children()
            .filter(|n| xml::is(*n, NS_OFFICE, "automatic-styles"))
        {
            for node in automatic
                .children()
                .filter(|n| xml::is(*n, NS_STYLE, "page-layout"))
            {
                budget::tick(1)?;
                let Some(name) = node.attribute((NS_STYLE, "name")) else {
                    continue;
                };
                let mut layout = default_layout.clone();
                read_layout(node, &mut layout)?;
                // A later declaration of one name replaces the earlier.
                match layout_index.get(name) {
                    Some(&index) => layouts[index] = layout,
                    None => {
                        layout_index.insert(name.to_string(), layouts.len());
                        layouts.push(layout);
                    }
                }
            }
        }
        let mut masters = Vec::new();
        let mut by_name = HashMap::new();
        let mut next_names: Vec<Option<&str>> = Vec::new();
        if let Some(styles) = styles_root
            .children()
            .find(|n| xml::is(*n, NS_OFFICE, "master-styles"))
        {
            for node in styles
                .children()
                .filter(|n| xml::is(*n, NS_STYLE, "master-page"))
            {
                budget::tick(1)?;
                let Some(name) = node.attribute((NS_STYLE, "name")) else {
                    continue;
                };
                by_name.entry(name.to_string()).or_insert(masters.len());
                next_names.push(
                    node.attribute((NS_STYLE, "next-style-name"))
                        .filter(|next| *next != name),
                );
                masters.push(Master {
                    name: name.to_string(),
                    layout: node
                        .attribute((NS_STYLE, "page-layout-name"))
                        .and_then(|layout| layout_index.get(layout).copied()),
                    next: None,
                    node,
                });
            }
        }
        for (master, next) in masters.iter_mut().zip(next_names) {
            master.next = next.and_then(|next| by_name.get(next).copied());
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
            .and_then(|m| m.layout)
            .map_or(&self.default_layout, |index| &self.layouts[index])
    }

    /// The master page whose layout the section's pages mostly use: a page
    /// style followed by another (`style:next-style-name`) lends its first
    /// page only.
    /// The layout of the master's body pages, through the resolved indices
    /// (the linked master's name is never hashed again).
    fn body_layout(&self, master: &str) -> &Layout {
        let Some(own) = self.master(master) else {
            return &self.default_layout;
        };
        let body = own.next.map_or(own, |index| &self.masters[index]);
        body.layout
            .map_or(&self.default_layout, |index| &self.layouts[index])
    }

    fn body_master<'s>(&'s self, master: &'s str) -> &'s str {
        self.master(master)
            .and_then(|m| m.next)
            .map_or(master, |index| self.masters[index].name.as_str())
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
        self.body_layout(master).columns.as_deref().cloned()
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
        let section = &self.body_layout(master).section;
        Geometry {
            text_width: (section.page_width - section.margin_left - section.margin_right).max(1.0),
            text_height: (section.page_height - section.margin_top - section.margin_bottom)
                .max(1.0),
            left: section.margin_left,
            right: section.margin_right,
        }
    }

    fn mirrored(&self) -> bool {
        self.layouts.iter().any(|layout| layout.usage == "mirrored")
    }

    /// Horizontal writing without a layout grid on the page style's body
    /// pages: the scope in which LibreOffice's fixed-line placement was
    /// measured.
    pub fn plain_horizontal(&self, master: &str) -> bool {
        let layout = self.body_layout(master);
        layout.text_direction.is_none() && layout.grid.is_none()
    }

    /// Horizontal writing on the page style's body pages, grid or not.
    pub fn horizontal(&self, master: &str) -> bool {
        self.body_layout(master).text_direction.is_none()
    }

    pub fn footnote_separator(&self, master: &str) -> Option<NoteSeparatorMark> {
        self.body_layout(master).footnote_separator
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
        budget::tick(1)?;
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
    read_background(node, layout)?;
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
        budget::tick(1)?;
        match (child.tag_name().namespace(), child.tag_name().name()) {
            (Some(NS_STYLE), "columns") => {
                let width =
                    (section.page_width - section.margin_left - section.margin_right).max(1.0);
                layout.columns = columns(child, width)?.map(Rc::new);
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
/// background image, from page layout properties. `layout` holds the
/// default page layout's (§16.8); what this element specifies replaces it,
/// property by property. (LibreOffice 26.8, measured, draws no background
/// from the default page layout at all; the specification is followed.)
fn read_background(node: roxmltree::Node<'_, '_>, layout: &mut Layout) -> Result<(), String> {
    let own = |ns: &str, name: &str| {
        (ns == NS_FO && name == "background-color")
            || (ns == NS_DRAW && (name.starts_with("fill") || name == "opacity"))
            || (ns == NS_STYLE && name == "repeat")
    };
    if node
        .attributes()
        .any(|a| own(a.namespace().unwrap_or(""), a.name()))
    {
        // (An element holds at most 256 attributes, each value bounded.)
        let mut attributes: HashMap<(String, String), String> = layout
            .background
            .as_ref()
            .and_then(|style| style.groups.get("graphic-properties"))
            .map(|properties| properties.attributes.clone())
            .unwrap_or_default();
        for attribute in node.attributes() {
            let namespace = attribute.namespace().unwrap_or("");
            if own(namespace, attribute.name()) {
                attributes.insert(
                    (namespace.to_string(), attribute.name().to_string()),
                    attribute.value().to_string(),
                );
            }
        }
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
    for child in node.children() {
        budget::tick(1)?;
        if xml::is(child, NS_STYLE, "background-image") {
            // The first such element is the property (§17.3).
            layout.background_image = child.attribute((NS_XLINK, "href")).map(|href| {
                Rc::new((
                    href.to_string(),
                    child.attribute((NS_STYLE, "repeat")) != Some("stretch"),
                ))
            });
            break;
        }
    }
    Ok(())
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

/// A `style:columns` element or one of its children as the reader meets
/// it: a parsed node (page layouts) or a captured style element (section
/// styles).
pub(super) trait ColumnsSource<'e>: Copy {
    fn value(self, namespace: &str, name: &str) -> Option<&'e str>;
    /// Whether this is the `style:` element `name`.
    fn is_style(self, name: &str) -> bool;
    /// Visits the child elements in order.
    fn each(self, visit: &mut dyn FnMut(Self) -> Result<(), String>) -> Result<(), String>;
}

impl<'a> ColumnsSource<'a> for roxmltree::Node<'a, '_> {
    fn value(self, namespace: &str, name: &str) -> Option<&'a str> {
        self.attribute((namespace, name))
    }
    fn is_style(self, name: &str) -> bool {
        xml::is(self, NS_STYLE, name)
    }
    fn each(self, visit: &mut dyn FnMut(Self) -> Result<(), String>) -> Result<(), String> {
        self.children()
            .filter(|n| n.is_element())
            .try_for_each(visit)
    }
}

impl<'e> ColumnsSource<'e> for &'e crate::style::Element {
    fn value(self, namespace: &str, name: &str) -> Option<&'e str> {
        self.attributes
            .get(&(namespace.to_string(), name.to_string()))
            .map(String::as_str)
    }
    fn is_style(self, name: &str) -> bool {
        self.namespace == NS_STYLE && self.name == name
    }
    fn each(self, visit: &mut dyn FnMut(Self) -> Result<(), String>) -> Result<(), String> {
        self.children.iter().try_for_each(visit)
    }
}

/// `style:columns` (§17.12): equal columns from `fo:column-count` and
/// `fo:column-gap`, or the `style:column` children's relative widths and
/// indents within `width`; a separator line (`style:column-sep`) maps to the
/// DOCX separator. More than `MAX_STYLE_COLUMNS` columns is beyond the
/// reader's limits; the count is validated before any column is read, and
/// no more than it are ever held.
pub(super) fn columns<'e, S: ColumnsSource<'e>>(
    element: S,
    width: f64,
) -> Result<Option<ColumnsSpec>, String> {
    let count = xml::bounded_count(
        element.value(NS_FO, "column-count"),
        MAX_STYLE_COLUMNS as u64,
        "column",
    )?
    .map_or(1, |count| count as usize);
    if count <= 1 {
        return Ok(None);
    }
    let gap = element
        .value(NS_FO, "column-gap")
        .and_then(|v| length_pt(v).ok())
        .unwrap_or(0.0);
    // One pass over the children: the separator, and the authored columns,
    // which are used only when there are exactly `count` of them (so at
    // most `count` are held: one more ends the collection).
    let mut sep = false;
    let mut authored: Vec<S> = Vec::new();
    let mut matching = true;
    element.each(&mut |child| {
        budget::tick(1)?;
        if child.is_style("column-sep") {
            sep |= child.value(NS_STYLE, "style") != Some("none");
        } else if child.is_style("column") && matching {
            if authored.len() == count {
                matching = false;
                authored.clear();
            } else {
                authored.push(child);
            }
        }
        Ok(())
    })?;
    if matching && authored.len() == count {
        let weights: Vec<f64> = authored
            .iter()
            .map(|c| {
                c.value(NS_STYLE, "rel-width")
                    .map(|v| v.trim_end_matches('*'))
                    .and_then(|v| v.parse::<f64>().ok())
                    .unwrap_or(1.0)
            })
            .collect();
        let total: f64 = weights.iter().sum();
        let indents: Vec<(f64, f64)> = authored
            .iter()
            .map(|c| {
                (
                    c.value(NS_FO, "start-indent")
                        .and_then(|v| length_pt(v).ok())
                        .unwrap_or(0.0),
                    c.value(NS_FO, "end-indent")
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
    pub fn read(styles: Option<roxmltree::Node<'_, '_>>) -> Result<Self, String> {
        let mut out = Self::default();
        let Some(styles) = styles else {
            return Ok(out);
        };
        for node in styles
            .children()
            .filter(|n| xml::is(*n, NS_TEXT, "notes-configuration"))
        {
            budget::tick(1)?;
            let config = NoteConfig {
                format: node
                    .attribute((NS_STYLE, "num-format"))
                    .and_then(|f| list::note_number_format(f)),
                // LibreOffice writes the offset from 1 (0 = start at 1).
                start: node
                    .attribute((NS_TEXT, "start-value"))
                    .and_then(|v| v.parse::<i64>().ok())
                    .and_then(|v| v.checked_add(1)),
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
        Ok(out)
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
        let base = self.region_stories_cached(&body_master)?;
        if body_master == master {
            self.headers.insert(master.to_string(), base.clone());
            return Ok(base);
        }
        // A first-page style shares its body style's stories: they are
        // built once (cached by that master) and copied here.
        budget::tick(1)?;
        let mut stories = MasterStories {
            headers: base.headers.clone(),
            footers: base.footers.clone(),
            title_page: base.title_page,
            has_left: base.has_left,
            shown: base.shown,
        };
        {
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
            stories.title_page = true;
        }
        let stories = Rc::new(stories);
        self.headers.insert(master.to_string(), stories.clone());
        Ok(stories)
    }

    /// One master page's own region stories, built once.
    fn region_stories_cached(&mut self, master: &str) -> Result<Rc<MasterStories>, String> {
        if let Some(stories) = self.region_cache.get(master) {
            return Ok(stories.clone());
        }
        let stories = Rc::new(self.region_stories(master)?);
        self.region_cache
            .insert(master.to_string(), stories.clone());
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
            let build = |reader: &mut Self,
                         node: roxmltree::Node<'a, 'a>|
             -> Result<HeaderFooter, String> {
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
                Ok(HeaderFooter { body })
            };
            // Lifetimes: master nodes live as long as the styles document.
            let default_node: roxmltree::Node<'a, 'a> = unsafe_node(default);
            let story = build(self, default_node)?;
            let left = shown(&format!("{kind}-left")).map(unsafe_node);
            let first = shown(&format!("{kind}-first")).map(unsafe_node);
            let even = match left {
                Some(left) => {
                    out.has_left = true;
                    Some(build(self, left)?)
                }
                None => None,
            };
            let first_story = match first {
                Some(first) => {
                    out.title_page = true;
                    Some(build(self, first)?)
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
        if let Some(background) = self.page_background(layout)? {
            // Each header story owns a copy of the background shape and its
            // paragraph.
            let paragraph =
                || BodyElement::Paragraph(Box::new(anchor_paragraph(vec![background.clone()])));
            out.headers
                .default
                .get_or_insert_with(HeaderFooter::default)
                .body
                .insert(0, paragraph());
            for story in [&mut out.headers.first, &mut out.headers.even]
                .into_iter()
                .flatten()
            {
                story.body.insert(0, paragraph());
            }
        }
        // A first page with only one region's first variant uses the other
        // region's ordinary story.
        if out.title_page {
            for set in [&mut out.headers, &mut out.footers] {
                if set.first.is_none() {
                    budget::tick(1)?;
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
        let layout = pages.body_layout(&start.master).clone();
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
                let base = self.default_font_size;
                ((width - base) * 4096.0).round()
            });
        }
        // Every section owns its copies of the master's stories (DOCX
        // sections resolve their header/footer references independently).
        // The limits are checked before each copy, so none is made past
        // them.
        budget::tick(1)?;
        let mut headers = stories.headers.clone();
        budget::tick(1)?;
        let mut footers = stories.footers.clone();
        if even_and_odd {
            for set in [&mut headers, &mut footers] {
                if set.even.is_none() {
                    budget::tick(1)?;
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
            budget::tick(1)?;
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
            budget::tick(1)?;
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
