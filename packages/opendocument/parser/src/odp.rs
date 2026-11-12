//! OpenDocument presentation reader (ODF 1.3: Part 3 §3.6
//! `office:presentation`, §9 drawing shapes, §10 frames, §16 styles) that
//! projects a packaged `.odp` / `.otp` directly into the PPTX slide model.
//!
//! Admission policy: every presentation feature is rendered; only corrupt,
//! encrypted or unsafe packages and packages over the resource limits are
//! rejected. Where the specification leaves a choice the reader follows
//! LibreOffice Impress, the reference producer and renderer.
//!
//! Slides are produced eagerly when the package is opened. Each slide model
//! lives only while it is built: it is serialized at once (the exact bytes
//! the WASM cursor returns per pull) and the deck retains those bytes.

use crate::error::{invalid, unsupported};
use crate::odp_geometry::{
    self as geo, enhanced_geometry, parse_points, parse_svg_path, parse_transform, parse_view_box,
    place, Affine, GeoPath, Geometry, Placed, Seg, ShapeFrame,
};
use crate::odp_media::{Media, Resource};
use crate::odp_paint::{self as paint, Line, Marker, PaintContext};
use crate::odp_style::{Nd, Props, Scope, Sheet, DRAWING_PAGE, GRAPHIC};
use crate::odp_text::{self as text, ObjectStyles, TextEnv, NS_PRESENTATION};
use crate::odp_units::{color, length_emu, percent};
use crate::package::{OdfPackage, SUPPORTED_VERSIONS};
use crate::xml::{self, NS_DRAW, NS_FO, NS_OFFICE, NS_STYLE, NS_SVG, NS_TABLE, NS_XLINK};
use ooxml_common::blip::{BlipEffect, SrcRect};
use ooxml_common::json_measurement::serialize_json_with_limit;
use ooxml_common::resource::{HardResourceLimitKind, HARD_MAX_PPTX_SLIDE_JSON_BYTES};
use pptx_model::{
    Fill, PathCmd, PathPaint, PictureElement, ShapeElement, Slide, SlideElement,
    SlideElementOrigin, SlideElementSource, SlideTheme, TextBody,
};
use std::collections::HashMap;
use std::rc::Rc;

pub const MEDIA_TYPES: [&str; 2] = [
    "application/vnd.oasis.opendocument.presentation",
    "application/vnd.oasis.opendocument.presentation-template",
];

/// Implementation bound on the elements of the whole deck (all slides,
/// master objects repeated per slide included), charged as each object's
/// elements are kept.
const MAX_DECK_ELEMENTS: usize = 2_000_000;
/// Implementation bound on the in-memory model of the slide being built:
/// each kept element's fixed size, serialized size and repeated records.
/// The deck retains only serialized slides (charged against
/// `odp_media::MAX_RETAINED_BYTES`).
const MAX_SLIDE_MODEL_BYTES: usize = 256 * 1024 * 1024;
/// Implementation bound on group nesting.
const MAX_GROUP_DEPTH: usize = 64;
/// Implementation bound on the drawing-object nodes the whole deck visits
/// (master objects once per slide that shows them, empty ones included).
const MAX_DECK_VISITS: usize = 64_000_000;
/// Implementation bound on the slides of a deck (the PPTX bootstrap limit).
const MAX_SLIDES: usize = 10_000;

pub struct OdpDeck {
    pub width: i64,
    pub height: i64,
    pub slides: Vec<OdpSlide>,
    pub resources: HashMap<String, Resource>,
    /// Converted resources (StarView metafiles as SVG), by package path.
    pub converted: HashMap<String, Vec<u8>>,
    pub hyperlink_colors: Option<(String, String)>,
    /// The slide models themselves, kept for unit tests only.
    #[cfg(test)]
    models: Vec<Slide>,
}

/// One read slide: the facts the bootstrap and tools need, and its model
/// serialized as the PPTX slide JSON the cursor returns.
pub struct OdpSlide {
    pub index: usize,
    pub part_name: Option<String>,
    pub hidden: bool,
    pub elements: usize,
    pub json: Vec<u8>,
}

/// Read the presentation. Every rejection happens here.
pub fn read(package: &OdfPackage) -> Result<OdpDeck, String> {
    let content_text = package.read_text("content.xml")?;
    let styles_text = if package.contains("styles.xml") {
        Some(package.read_text("styles.xml")?)
    } else {
        None
    };
    let content = xml::parse(&content_text)?;
    let styles = match &styles_text {
        Some(t) => Some(xml::parse(t)?),
        None => None,
    };
    let root = content.root_element();
    if !xml::is(root, NS_OFFICE, "document-content") {
        return Err(invalid("content.xml is not office:document-content"));
    }
    let version = root.attribute((NS_OFFICE, "version"));
    if !version.is_some_and(|v| SUPPORTED_VERSIONS.contains(&v)) {
        return Err(unsupported("odf-version"));
    }
    let mut sheet = Sheet::default();
    if let Some(styles) = &styles {
        sheet.add_document(styles.root_element(), Scope::Styles)?;
    }
    sheet.add_document(root, Scope::Content)?;
    let presentation = root
        .children()
        .find(|n| xml::is(*n, NS_OFFICE, "body"))
        .and_then(|body| {
            body.children()
                .find(|n| xml::is(*n, NS_OFFICE, "presentation"))
        })
        .ok_or_else(|| invalid("the document has no office:presentation body"))?;
    let mut media = Media::new(package);
    // Header/footer declarations: one shared copy each, charged once.
    let mut decls: HashMap<(&str, &str), Rc<str>> = HashMap::new();
    for child in presentation.children().filter(|n| n.is_element()) {
        if child.tag_name().namespace() == Some(NS_PRESENTATION) {
            let kind = child.tag_name().name();
            if matches!(kind, "footer-decl" | "header-decl" | "date-time-decl") {
                if let Some(name) = child.attribute((NS_PRESENTATION, "name")) {
                    let value = text::plain_text_inline(child)?;
                    media.charge(value.len())?;
                    decls.insert((kind, name), value);
                }
            }
        }
    }
    let pages: Vec<Nd> = presentation
        .children()
        .filter(|n| xml::is(*n, NS_DRAW, "page"))
        .collect();
    if pages.len() > MAX_SLIDES {
        return Err("OpenDocument presentation exceeds the slide limit".to_string());
    }
    let page_names: HashMap<String, usize> = pages
        .iter()
        .enumerate()
        .filter_map(|(i, p)| p.attribute((NS_DRAW, "name")).map(|n| (n.to_string(), i)))
        .collect();
    // Slide size: the page layout of the first slide's master page (§16.5).
    let first_master = pages
        .first()
        .and_then(|p| p.attribute((NS_DRAW, "master-page-name")))
        .and_then(|name| sheet.master_pages.get(name).copied())
        .or_else(|| {
            sheet
                .master_order
                .first()
                .and_then(|n| sheet.master_pages.get(n).copied())
        });
    let (width, height) = first_master
        .and_then(|m| m.attribute((NS_STYLE, "page-layout-name")))
        .and_then(|name| sheet.page_layout(name))
        .and_then(|layout| {
            layout
                .children()
                .find(|c| xml::is(*c, NS_STYLE, "page-layout-properties"))
        })
        .map(|props| {
            (
                props.attribute((NS_FO, "page-width")).and_then(length_emu),
                props.attribute((NS_FO, "page-height")).and_then(length_emu),
            )
        })
        .and_then(|(w, h)| Some((w?, h?)))
        .unwrap_or((28.0 * 360_000.0, 21.0 * 360_000.0));

    let mut budget = MAX_DECK_ELEMENTS;
    let mut visits = MAX_DECK_VISITS;
    let text_work = std::cell::Cell::new(text::MAX_DECK_TEXT_WORK);
    let mut master_frames = HashMap::new();
    let mut marker_paths = HashMap::new();
    // Masters whose theme was already searched (one scan each).
    let mut themes_searched: Vec<roxmltree::NodeId> = Vec::new();
    let mut charts: HashMap<String, Option<pptx_model::ChartModel>> = HashMap::new();
    let mut slides = Vec::with_capacity(pages.len());
    #[cfg(test)]
    let mut models = Vec::new();
    let mut hyperlink_colors = None;
    for (index, page) in pages.iter().enumerate() {
        let master = page
            .attribute((NS_DRAW, "master-page-name"))
            .and_then(|name| sheet.master_pages.get(name).copied())
            .or(first_master);
        if let Some(m) = master.filter(|_| hyperlink_colors.is_none()) {
            if !themes_searched.contains(&m.id()) {
                themes_searched.push(m.id());
                hyperlink_colors = theme_hyperlinks(m);
            }
        }
        let mut reader = SlideReader {
            sheet: &sheet,
            media: &mut media,
            index,
            count: pages.len(),
            page_names: &page_names,
            budget: &mut budget,
            visits: &mut visits,
            text_work: &text_work,
            master_frames: &mut master_frames,
            charts: &mut charts,
            slide_model: MAX_SLIDE_MODEL_BYTES,
            marker_paths: &mut marker_paths,
            next_id: 1,
            page_size: (width, height),
        };
        let slide = reader.slide(*page, master, &decls)?;
        let json = serialize_json_with_limit(
            &slide,
            None,
            HardResourceLimitKind::PptxSlideJsonBytes,
            None,
            HARD_MAX_PPTX_SLIDE_JSON_BYTES,
            "OpenDocument slide exceeds the PPTX slide JSON ceiling",
        )?;
        let kept = OdpSlide {
            index: slide.index,
            part_name: slide.part_name.clone(),
            hidden: slide.hidden,
            elements: slide.elements.len(),
            json,
        };
        #[cfg(test)]
        models.push(slide.clone());
        drop(slide);
        media.charge(
            std::mem::size_of::<OdpSlide>()
                .saturating_add(kept.json.len())
                .saturating_add(kept.part_name.as_ref().map_or(0, String::len)),
        )?;
        slides.push(kept);
    }
    let (resources, converted) = media.into_admitted();
    Ok(OdpDeck {
        width: width.round() as i64,
        height: height.round() as i64,
        slides,
        resources,
        converted,
        hyperlink_colors,
        #[cfg(test)]
        models,
    })
}

/// LibreOffice's master-page theme (`loext:theme`): hyperlink colours.
fn theme_hyperlinks(master: Nd<'_>) -> Option<(String, String)> {
    let colors = master
        .descendants()
        .find(|n| xml::is(*n, crate::xml::NS_LOEXT, "theme-colors"))?;
    let get = |name: &str| {
        colors
            .children()
            .find(|c| c.attribute((crate::xml::NS_LOEXT, "name")) == Some(name))
            .and_then(|c| c.attribute((crate::xml::NS_LOEXT, "color")))
            .and_then(color)
    };
    Some((get("hyperlink")?, get("followed-hyperlink")?))
}

struct SlideReader<'a, 'r, 'p> {
    sheet: &'a Sheet<'a>,
    media: &'r mut Media<'p>,
    index: usize,
    count: usize,
    page_names: &'r HashMap<String, usize>,
    budget: &'r mut usize,
    /// Visits left of [`MAX_DECK_VISITS`].
    visits: &'r mut usize,
    /// Text work left of `odp_text::MAX_DECK_TEXT_WORK`.
    text_work: &'r std::cell::Cell<usize>,
    /// Each master's presentation frames by class, for placeholders that
    /// take their placement from the master (indexed once per master).
    master_frames: &'r mut HashMap<roxmltree::NodeId, HashMap<&'a str, Nd<'a>>>,
    /// Chart objects read so far by package path (`None`: not a chart),
    /// so a chart a master shows on every slide is read once.
    charts: &'r mut HashMap<String, Option<pptx_model::ChartModel>>,
    /// Model bytes left for this slide (see [`MAX_SLIDE_MODEL_BYTES`]).
    slide_model: usize,
    /// Parsed marker paths by their definition's `svg:d` (address, length).
    marker_paths: &'r mut HashMap<(usize, usize), Option<Vec<Seg>>>,
    next_id: u32,
    /// Slide size in EMU (page backgrounds are laid out on it).
    page_size: (f64, f64),
}

/// Placeholder classes of a master page that are layout templates, not
/// drawn objects (§19.394 `presentation:class`).
fn is_master_template(class: &str) -> bool {
    !matches!(
        class,
        "header" | "footer" | "date-time" | "page-number" | ""
    )
}

/// Per-slide context for one drawn object.
struct ObjectContext<'a> {
    scope: Scope,
    origin: SlideElementOrigin,
    base: &'a str,
    master_name: Option<&'a str>,
    background: Option<Fill>,
    background_color: Option<String>,
    footer: Option<Rc<str>>,
    header: Option<Rc<str>>,
    date_time: Option<Rc<str>>,
    link: Option<String>,
    /// The slide's master page, for placeholders that take their placement
    /// from it (§19.431); `None` for the master's own objects.
    master: Option<Nd<'a>>,
}

impl<'a, 'r, 'p> SlideReader<'a, 'r, 'p> {
    fn slide(
        &mut self,
        page: Nd<'a>,
        master: Option<Nd<'a>>,
        decls: &HashMap<(&str, &str), Rc<str>>,
    ) -> Result<Slide, String> {
        let sheet = self.sheet;
        let page_styles = page
            .attribute((NS_DRAW, "style-name"))
            .map(|n| sheet.chain(Scope::Content, "drawing-page", n))
            .transpose()?
            .unwrap_or_default();
        let master_styles = master
            .and_then(|m| m.attribute((NS_DRAW, "style-name")))
            .map(|n| sheet.chain(Scope::Styles, "drawing-page", n))
            .transpose()?
            .unwrap_or_default();
        let mut page_props = Props::new();
        page_props.extend_styles(&page_styles, DRAWING_PAGE);
        let mut master_props = Props::new();
        master_props.extend_styles(&master_styles, DRAWING_PAGE);
        self.visit(
            crate::odp_style::props_cost(&page_props)
                .saturating_add(crate::odp_style::props_cost(&master_props)),
        )?;
        let flag = |name: &str| page_props.get(NS_PRESENTATION, name) != Some("false");
        let hidden = page_props.get(NS_PRESENTATION, "visibility") == Some("hidden");

        // Background (§20.4): the slide's own fill, else the master's, unless
        // presentation:background-visible hides the master background.
        let area = self.page_size;
        let own_fill = page_props.get(NS_DRAW, "fill");
        let mut background = None;
        if matches!(own_fill, Some(kind) if kind != "none") {
            let mut ctx = PaintContext {
                sheet,
                media: self.media,
                base: "content.xml",
                slide_background: None,
            };
            background = paint::page_fill(&page_props, &mut ctx, area);
        }
        if background.is_none() && flag("background-visible") {
            let mut ctx = PaintContext {
                sheet,
                media: self.media,
                base: "styles.xml",
                slide_background: None,
            };
            background = paint::page_fill(&master_props, &mut ctx, area);
        }
        let background_color = match &background {
            Some(Fill::Solid { color }) => Some(color[..6.min(color.len())].to_string()),
            Some(Fill::Gradient { stops, .. }) => stops
                .first()
                .map(|s| s.color[..6.min(s.color.len())].to_string()),
            _ => None,
        };
        if let Some(fill) = &background {
            self.slide_model = self
                .slide_model
                .checked_sub(text::record_bytes(fill)?)
                .ok_or_else(|| "OpenDocument slide exceeds the slide model limit".to_string())?;
        }
        self.media.status()?;
        let decl = |kind: &str, attr: &str| -> Option<Rc<str>> {
            let name = page.attribute((NS_PRESENTATION, attr))?;
            decls.get(&(kind, name)).cloned()
        };
        let footer = decl("footer-decl", "use-footer-name");
        let header = decl("header-decl", "use-header-name");
        let date_time = decl("date-time-decl", "use-date-time-name");
        let master_name = master.and_then(|m| m.attribute((NS_STYLE, "name")));

        let mut elements = Vec::new();
        let mut sources = Vec::new();
        if let Some(master) = master {
            if flag("background-objects-visible") {
                let ctx = ObjectContext {
                    scope: Scope::Styles,
                    origin: SlideElementOrigin::Master,
                    base: "styles.xml",
                    master_name,
                    background: background.clone(),
                    background_color: background_color.clone(),
                    footer: footer.clone(),
                    header: header.clone(),
                    date_time: date_time.clone(),
                    link: None,
                    master: None,
                };
                // The display flags, resolved once for the slide.
                let shows = [
                    ("footer", flag("display-footer")),
                    ("header", flag("display-header")),
                    ("date-time", flag("display-date-time")),
                    ("page-number", flag("display-page-number")),
                ];
                let children: Vec<Nd<'a>> = master.children().filter(|n| n.is_element()).collect();
                self.visit(children.len())?;
                let shown: Vec<Nd<'a>> = children
                    .into_iter()
                    .filter(|child| {
                        // A hyperlinked shape is classed by the shape it wraps.
                        let class = wrapped_shape(*child)
                            .attribute((NS_PRESENTATION, "class"))
                            .unwrap_or("");
                        !is_master_template(class)
                            && shows
                                .iter()
                                .find(|(kind, _)| *kind == class)
                                .is_none_or(|(_, shown)| *shown)
                    })
                    .collect();
                for child in stacking_order(shown) {
                    self.object(child, &ctx, &mut elements, 0)?;
                }
            }
        }
        let master_count = elements.len();
        sources.extend((0..master_count).map(|_| SlideElementSource {
            origin: SlideElementOrigin::Master,
        }));
        let ctx = ObjectContext {
            scope: Scope::Content,
            origin: SlideElementOrigin::Slide,
            base: "content.xml",
            master_name,
            background: background.clone(),
            background_color,
            footer,
            header,
            date_time,
            link: None,
            master,
        };
        let mut notes = None;
        let mut objects = Vec::new();
        for child in page.children().filter(|n| n.is_element()) {
            if xml::is(child, NS_PRESENTATION, "notes") {
                let text = text::notes_text(
                    child
                        .children()
                        .filter(|n| n.attribute((NS_PRESENTATION, "class")) == Some("notes")),
                )?;
                if !text.trim().is_empty() {
                    notes = Some(text);
                }
                continue;
            }
            objects.push(child);
        }
        for child in stacking_order(objects) {
            self.object(child, &ctx, &mut elements, 0)?;
        }
        sources.extend((master_count..elements.len()).map(|_| SlideElementSource {
            origin: SlideElementOrigin::Slide,
        }));
        Ok(Slide {
            theme: Some(SlideTheme {
                default_text_color: Some("000000".into()),
                major_font: None,
                minor_font: None,
                hlink_color: None,
            }),
            index: self.index,
            slide_number: self.index + 1,
            part_name: Some(text::slide_part_name(self.index)),
            background: background.or(Some(Fill::Solid {
                color: "FFFFFF".into(),
            })),
            elements,
            element_sources: sources,
            notes,
            comments: Vec::new(),
            hidden,
            parse_error: None,
        })
    }

    fn id(&mut self) -> String {
        let id = self.next_id;
        self.next_id += 1;
        id.to_string()
    }

    /// The style chain of a drawing object: `presentation:style-name`
    /// (presentation family) or `draw:style-name` (graphic family), then
    /// the default graphic style.
    fn object_styles(
        &self,
        node: Nd<'a>,
        ctx: &ObjectContext<'a>,
    ) -> Result<ObjectStyles<'a>, String> {
        let sheet = self.sheet;
        let (family, name) = match node.attribute((NS_PRESENTATION, "style-name")) {
            Some(name) => ("presentation", Some(name)),
            None => ("graphic", node.attribute((NS_DRAW, "style-name"))),
        };
        let chain = match name {
            Some(name) => sheet.chain(ctx.scope, family, name)?,
            None => Vec::new(),
        };
        let automatic = name.and_then(|n| {
            let found = sheet.find(ctx.scope, family, n)?;
            (sheet.common(family, n) != Some(found)).then_some(found)
        });
        let class = node.attribute((NS_PRESENTATION, "class"));
        let text_style = node
            .attribute((NS_DRAW, "text-style-name"))
            .map(|n| sheet.chain(ctx.scope, "paragraph", n))
            .transpose()?
            .unwrap_or_default();
        Ok(ObjectStyles {
            chain,
            automatic,
            outline_master: if class == Some("outline") {
                ctx.master_name
            } else {
                None
            },
            text_style,
            defaults: sheet.default_style("graphic").into_iter().collect(),
        })
    }

    /// The object's graphic properties; reading them is charged as visits.
    fn graphic_props(&mut self, styles: &ObjectStyles<'a>) -> Result<Props<'a>, String> {
        let mut props = Props::new();
        props.extend_styles(&styles.chain, GRAPHIC);
        props.extend_styles(&styles.defaults, GRAPHIC);
        self.visit(crate::odp_style::props_cost(&props))?;
        Ok(props)
    }

    fn object(
        &mut self,
        node: Nd<'a>,
        ctx: &ObjectContext<'a>,
        out: &mut Vec<SlideElement>,
        depth: usize,
    ) -> Result<(), String> {
        if depth > MAX_GROUP_DEPTH {
            return Err(invalid("drawing groups are nested too deeply"));
        }
        self.visit(1)?;
        if node.tag_name().namespace() != Some(NS_DRAW) {
            return Ok(());
        }
        // Empty presentation placeholders are not shown in a slide show.
        if node.attribute((NS_PRESENTATION, "placeholder")) == Some("true") {
            return Ok(());
        }
        match node.tag_name().name() {
            "g" => {
                for child in node.children().filter(|n| n.is_element()) {
                    self.object(child, ctx, out, depth + 1)?;
                }
            }
            "a" => {
                // The target is copied for this traversal: charged by size.
                let href = node.attribute((NS_XLINK, "href"));
                self.visit(href.map_or(0, |h| h.len() / 64))?;
                let link = href.map(str::to_string);
                let inner = ObjectContext {
                    link,
                    ..clone_ctx(ctx)
                };
                for child in node.children().filter(|n| n.is_element()) {
                    self.object(child, &inner, out, depth + 1)?;
                }
            }
            "frame" => {
                // A drawn object costs work in proportion to its markup
                // (attributes, alternatives, text), each time it is drawn.
                self.visit(node.range().len() / 64)?;
                let before = out.len();
                self.frame(node, ctx, out)?;
                self.retain(&out[before..])?;
            }
            "custom-shape" | "rect" | "ellipse" | "circle" | "line" | "polyline" | "polygon"
            | "regular-polygon" | "path" | "connector" | "caption" | "measure" => {
                self.visit(node.range().len() / 64)?;
                let before = out.len();
                self.shape(node, ctx, out)?;
                self.retain(&out[before..])?;
            }
            _ => {}
        }
        Ok(())
    }

    /// Charge `count` object visits against the deck's bound.
    fn visit(&mut self, count: usize) -> Result<(), String> {
        *self.visits = self
            .visits
            .checked_sub(count)
            .ok_or_else(|| "OpenDocument presentation exceeds the object limit".to_string())?;
        Ok(())
    }

    /// Charge the elements one object produced against the deck's element
    /// and retained-size bounds as they are kept (the object's own content
    /// was bounded while it was read: text per object, geometry per shape,
    /// cells per table, points per chart).
    fn retain(&mut self, elements: &[SlideElement]) -> Result<(), String> {
        *self.budget = self
            .budget
            .checked_sub(elements.len())
            .ok_or_else(|| "OpenDocument presentation exceeds the element limit".to_string())?;
        for element in elements {
            self.slide_model = self
                .slide_model
                .checked_sub(model_bytes(element)?)
                .ok_or_else(|| "OpenDocument slide exceeds the slide model limit".to_string())?;
        }
        // Resources generated for the object (hatches, metafiles).
        self.media.status()
    }

    /// The node whose size and position a frame takes: with
    /// `presentation:user-transformed="false"` (§19.431) the master page's
    /// presentation shape of the same class (the first one), else the frame
    /// itself.
    fn placement_node(&mut self, node: Nd<'a>, ctx: &ObjectContext<'a>) -> Result<Nd<'a>, String> {
        if node.attribute((NS_PRESENTATION, "user-transformed")) != Some("false") {
            return Ok(node);
        }
        let (Some(class), Some(master)) = (node.attribute((NS_PRESENTATION, "class")), ctx.master)
        else {
            return Ok(node);
        };
        if !self.master_frames.contains_key(&master.id()) {
            let mut frames = HashMap::new();
            let mut visited = 0usize;
            for child in master.children().filter(|c| c.is_element()) {
                visited += 1;
                let shape = wrapped_shape(child);
                if !xml::is(shape, NS_DRAW, "frame") {
                    continue;
                }
                if let Some(c) = shape.attribute((NS_PRESENTATION, "class")) {
                    frames.entry(c).or_insert(shape);
                }
            }
            self.visit(visited)?;
            self.master_frames.insert(master.id(), frames);
        }
        Ok(self.master_frames[&master.id()]
            .get(class)
            .copied()
            .unwrap_or(node))
    }

    /// The placed box of an object: svg:x/y/width/height then draw:transform.
    fn placement(node: Nd<'_>, w: f64, h: f64) -> Placed {
        let x = node
            .attribute((NS_SVG, "x"))
            .and_then(length_emu)
            .unwrap_or(0.0);
        let y = node
            .attribute((NS_SVG, "y"))
            .and_then(length_emu)
            .unwrap_or(0.0);
        let mut m = Affine::translate(x, y);
        if let Some(t) = node
            .attribute((NS_DRAW, "transform"))
            .and_then(parse_transform)
        {
            m = m.then(&t);
        }
        place(w, h, &m)
    }

    fn size(node: Nd<'_>) -> (f64, f64) {
        (
            node.attribute((NS_SVG, "width"))
                .and_then(length_emu)
                .unwrap_or(0.0)
                .max(0.0),
            node.attribute((NS_SVG, "height"))
                .and_then(length_emu)
                .unwrap_or(0.0)
                .max(0.0),
        )
    }

    fn shape(
        &mut self,
        node: Nd<'a>,
        ctx: &ObjectContext<'a>,
        out: &mut Vec<SlideElement>,
    ) -> Result<(), String> {
        let styles = self.object_styles(node, ctx)?;
        let graphic = self.graphic_props(&styles)?;
        let kind = node.tag_name().name();
        let line = paint::line(&graphic, self.sheet);
        let has_stroke = line.stroke.is_some();
        // Geometry in shape-local EMU and the placed box.
        let (mut w, mut h) = Self::size(node);
        let mut flip_h = false;
        let mut flip_v = false;
        let mut absolute = false;
        let mut placed: Option<Placed> = None;
        let mut geometry = match kind {
            "custom-shape" => {
                let fill_kind = graphic.get(NS_DRAW, "fill").unwrap_or("none");
                let frame = ShapeFrame {
                    width: w,
                    height: h,
                    has_fill: fill_kind != "none",
                    has_stroke,
                };
                match node
                    .children()
                    .find(|c| xml::is(*c, NS_DRAW, "enhanced-geometry"))
                {
                    Some(eg) => {
                        flip_h = eg.attribute((NS_DRAW, "mirror-horizontal")) == Some("true");
                        flip_v = eg.attribute((NS_DRAW, "mirror-vertical")) == Some("true");
                        enhanced_geometry(eg, &frame)?
                    }
                    None => rect_geometry(w, h, 0.0),
                }
            }
            "rect" => {
                let r = node
                    .attribute((NS_DRAW, "corner-radius"))
                    .and_then(length_emu)
                    .unwrap_or(0.0);
                rect_geometry(w, h, r)
            }
            "ellipse" | "circle" => ellipse_geometry(node, w, h),
            "regular-polygon" => regular_polygon(node, w, h),
            "polyline" | "polygon" => {
                let points = node
                    .attribute((NS_DRAW, "points"))
                    .and_then(parse_points)
                    .unwrap_or_default();
                let vb = node.attribute((NS_SVG, "viewBox")).and_then(parse_view_box);
                let mut segs = Vec::new();
                for (i, (x, y)) in points.iter().enumerate() {
                    let (x, y) = view_map(vb, w, h, *x, *y);
                    segs.push(if i == 0 {
                        Seg::Move(x, y)
                    } else {
                        Seg::Line(x, y)
                    });
                }
                if kind == "polygon" && !segs.is_empty() {
                    segs.push(Seg::Close);
                }
                single(segs, kind == "polygon")
            }
            "path" => {
                let segs = node
                    .attribute((NS_SVG, "d"))
                    .and_then(parse_svg_path)
                    .unwrap_or_default();
                let vb = node.attribute((NS_SVG, "viewBox")).and_then(parse_view_box);
                let segs = segs
                    .into_iter()
                    .map(|s| geo::map_seg(s, |x, y| view_map(vb, w, h, x, y)))
                    .collect();
                single(segs, true)
            }
            "line" => {
                absolute = true;
                let p = |a: &str| {
                    node.attribute((NS_SVG, a))
                        .and_then(length_emu)
                        .unwrap_or(0.0)
                };
                single(
                    vec![Seg::Move(p("x1"), p("y1")), Seg::Line(p("x2"), p("y2"))],
                    false,
                )
            }
            "connector" => {
                absolute = true;
                let p = |a: &str| {
                    node.attribute((NS_SVG, a))
                        .and_then(length_emu)
                        .unwrap_or(0.0)
                };
                // LibreOffice writes the routed path in page coordinates
                // (1/100 mm); without one the connector is a straight line.
                match node.attribute((NS_SVG, "d")).and_then(parse_svg_path) {
                    Some(segs) if !segs.is_empty() => single(
                        segs.into_iter()
                            .map(|s| geo::map_seg(s, |x, y| (x * 360.0, y * 360.0)))
                            .collect(),
                        false,
                    ),
                    _ => single(
                        vec![Seg::Move(p("x1"), p("y1")), Seg::Line(p("x2"), p("y2"))],
                        false,
                    ),
                }
            }
            "caption" => {
                // The caption box; the leader line runs to draw:caption-point.
                let mut g = rect_geometry(
                    w,
                    h,
                    node.attribute((NS_DRAW, "corner-radius"))
                        .and_then(length_emu)
                        .unwrap_or(0.0),
                );
                if let (Some(cx), Some(cy)) = (
                    node.attribute((NS_DRAW, "caption-point-x"))
                        .and_then(length_emu),
                    node.attribute((NS_DRAW, "caption-point-y"))
                        .and_then(length_emu),
                ) {
                    g.paths.push(GeoPath {
                        segs: vec![Seg::Move(w / 2.0, h / 2.0), Seg::Line(cx, cy)],
                        fill: false,
                        stroke: true,
                        fill_mode: None,
                    });
                }
                g
            }
            "measure" => {
                absolute = true;
                let p = |a: &str| {
                    node.attribute((NS_SVG, a))
                        .and_then(length_emu)
                        .unwrap_or(0.0)
                };
                single(
                    vec![Seg::Move(p("x1"), p("y1")), Seg::Line(p("x2"), p("y2"))],
                    false,
                )
            }
            _ => return Ok(()),
        };
        if absolute {
            // Geometry is in page coordinates: its bounds are the box.
            let Some((x0, y0, x1, y1)) = geometry.bounds() else {
                return Ok(());
            };
            geometry.translate(-x0, -y0);
            w = x1 - x0;
            h = y1 - y0;
            let m = node
                .attribute((NS_DRAW, "transform"))
                .and_then(parse_transform)
                .unwrap_or(Affine::IDENTITY);
            placed = Some(place(w, h, &Affine::translate(x0, y0).then(&m)));
        }
        let placed = placed.unwrap_or_else(|| Self::placement(node, w, h));
        flip_v ^= placed.flip_v;
        // Line ends: markers drawn as filled shapes; the line is shortened
        // under them (LibreOffice's arrow docking).
        let markers = self.markers(&mut geometry, &line, w, h, &placed, flip_h, flip_v);
        let mut paint_ctx = PaintContext {
            sheet: self.sheet,
            media: self.media,
            base: ctx.base,
            slide_background: ctx.background.as_ref(),
        };
        // Fontwork (§19.164 draw:text-path): the geometry is the path the
        // text follows, not a painted outline. The LibreOffice fontwork
        // shapes are the DrawingML text warps they were imported from.
        let warp = node
            .children()
            .find(|c| xml::is(*c, NS_DRAW, "enhanced-geometry"))
            .filter(|eg| eg.attribute((NS_DRAW, "text-path")) == Some("true"))
            .map(|eg| fontwork_warp(eg.attribute((NS_DRAW, "type")).unwrap_or("")));
        let mut fill = paint::fill(&graphic, &mut paint_ctx, (placed.w, placed.h));
        let mut stroke = line.stroke;
        if warp.is_some() {
            fill = None;
            stroke = None;
        }
        let background = solid_color(&fill).or_else(|| ctx.background_color.clone());
        let mut text_body = self.text(node, ctx, &styles, &graphic, background)?;
        if let (Some(body), Some(preset)) = (&mut text_body, warp) {
            body.text_warp = Some(pptx_model::TextWarp {
                preset: preset.to_string(),
                adj: Vec::new(),
            });
        }
        let text_rect = geometry.text_area.map(|(l, t, tw, th)| {
            // The renderer lays text out in the unmirrored frame and turns it
            // by 180° for a vertical mirror; place the area so the displayed
            // rectangle is the mirrored geometry's.
            let rx = if flip_h != flip_v { w - l - tw } else { l };
            pptx_model::TextRect {
                x: (placed.x + rx * placed.w / w.max(1e-9)).round() as i64,
                y: (placed.y + t * placed.h / h.max(1e-9)).round() as i64,
                width: (tw * placed.w / w.max(1e-9)).round() as i64,
                height: (th * placed.h / h.max(1e-9)).round() as i64,
            }
        });
        let (cust_geom, cust_geom_paint) = normalize(&geometry, w, h);
        // An extruded custom shape is drawn through the 3-D scene; its 2-D
        // rotation is part of the camera revolution.
        let extrusion = if kind == "custom-shape" {
            node.children()
                .find(|c| xml::is(*c, NS_DRAW, "enhanced-geometry"))
                .and_then(|eg| crate::odp_extrusion::scene(eg, &graphic, -placed.rot))
        } else {
            None
        };
        let (scene3d, sp3d) = match extrusion {
            Some((scene, sp)) => (Some(scene), Some(sp)),
            None => (None, None),
        };
        let rotation = if scene3d.is_some() { 0.0 } else { placed.rot };
        let element = ShapeElement {
            math_fallback: None,
            x: placed.x.round() as i64,
            y: placed.y.round() as i64,
            width: placed.w.round() as i64,
            height: placed.h.round() as i64,
            rotation,
            flip_h,
            flip_v,
            geometry: "custGeom".into(),
            fill,
            stroke,
            text_body,
            default_text_color: None,
            cust_geom: Some(cust_geom),
            cust_geom_paint,
            adj: None,
            adj2: None,
            adj3: None,
            adj4: None,
            adj5: None,
            adj6: None,
            adj7: None,
            adj8: None,
            shadow: paint::shadow(&graphic),
            inner_shadow: None,
            glow: paint::glow(&graphic),
            soft_edge: paint::soft_edge(&graphic),
            reflection: None,
            id: Some(self.id()),
            name: node.attribute((NS_DRAW, "name")).map(str::to_string),
            hyperlink: ctx.link.clone(),
            hyperlink_action: None,
            placeholder_type: placeholder_type(node),
            placeholder_idx: None,
            text_rect,
            scene3d,
            sp3d,
        };
        out.push(SlideElement::Shape(element));
        out.extend(markers);
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn markers(
        &mut self,
        geometry: &mut Geometry,
        line: &Line<'a>,
        w: f64,
        h: f64,
        placed: &Placed,
        flip_h: bool,
        flip_v: bool,
    ) -> Vec<SlideElement> {
        let mut out = Vec::new();
        let to_page = |x: f64, y: f64| -> (f64, f64) {
            let lx = if flip_h { w - x } else { x } * placed.w / w.max(1e-9);
            let ly = if flip_v { h - y } else { y } * placed.h / h.max(1e-9);
            let (cx, cy) = (placed.w / 2.0, placed.h / 2.0);
            let (s, c) = placed.rot.to_radians().sin_cos();
            let (dx, dy) = (lx - cx, ly - cy);
            (
                placed.x + cx + dx * c - dy * s,
                placed.y + cy + dx * s + dy * c,
            )
        };
        let ends: [(Option<&Marker>, bool); 2] = [
            (line.start_marker.as_ref(), true),
            (line.end_marker.as_ref(), false),
        ];
        for (marker, at_start) in ends {
            let Some(marker) = marker else { continue };
            let subpath = if at_start {
                geometry
                    .paths
                    .iter_mut()
                    .find(|p| p.stroke && !p.segs.is_empty())
            } else {
                geometry
                    .paths
                    .iter_mut()
                    .rev()
                    .find(|p| p.stroke && !p.segs.is_empty())
            };
            let Some(path) = subpath else { continue };
            // The open subpath at that end.
            let (start, end) = subpath_range(&path.segs, at_start);
            let segs = &mut path.segs[start..end];
            if segs.iter().any(|s| matches!(s, Seg::Close)) {
                continue;
            }
            let tangent = if at_start {
                geo::start_tangent(segs)
            } else {
                geo::end_tangent(segs)
            };
            let Some((tip_local, outward_local)) = tangent else {
                continue;
            };
            // Work in page space so a non-uniform placement keeps the marker's
            // own proportions.
            let tip = to_page(tip_local.0, tip_local.1);
            let probe = to_page(
                tip_local.0 - outward_local.0 * 1000.0,
                tip_local.1 - outward_local.1 * 1000.0,
            );
            let (ix, iy) = (probe.0 - tip.0, probe.1 - tip.1);
            let len = ix.hypot(iy);
            if len <= 0.0 {
                continue;
            }
            let inward = (ix / len, iy / len);
            let (vx, vy, vw, vh) = marker.view_box;
            if vw <= 0.0 || vh <= 0.0 {
                continue;
            }
            let scale = marker.width / vw;
            let marker_len = vh * scale;
            let docking = if marker.center { 0.5 } else { 0.0 };
            // Tip offset: a centred marker is centred on the line end.
            let tip_page = (
                tip.0 - inward.0 * marker_len * docking,
                tip.1 - inward.1 * marker_len * docking,
            );
            let perp = (inward.1, -inward.0);
            // Parsed once per marker definition (kept for the deck).
            let key = (marker.d.as_ptr() as usize, marker.d.len());
            if !self.marker_paths.contains_key(&key) {
                let parsed = parse_svg_path(marker.d);
                if let Some(segs) = &parsed {
                    // Past the retained limit the deck fails at `retain`.
                    let bytes = segs.len().saturating_mul(std::mem::size_of::<Seg>());
                    if self.media.charge(bytes).is_err() {
                        continue;
                    }
                }
                self.marker_paths.insert(key, parsed);
            }
            let Some(segs_marker) = self.marker_paths.get(&key).and_then(Option::as_ref) else {
                continue;
            };
            let mapped: Vec<Seg> = segs_marker
                .iter()
                .cloned()
                .map(|s| {
                    geo::map_seg(s, |u, v| {
                        let du = (u - vx - vw / 2.0) * scale;
                        let dv = (v - vy) * scale;
                        (
                            tip_page.0 + inward.0 * dv + perp.0 * du,
                            tip_page.1 + inward.1 * dv + perp.1 * du,
                        )
                    })
                })
                .collect();
            // Shorten the line under the marker (in local units).
            let local_scale = if w > 0.0 && placed.w > 0.0 {
                placed.w / w
            } else {
                1.0
            };
            let consumed = marker_len * (1.0 - docking) / local_scale;
            if at_start {
                geo::trim_start(segs, consumed);
            } else {
                geo::trim_end(segs, consumed);
            }
            let mut g = Geometry {
                paths: vec![GeoPath {
                    segs: mapped,
                    fill: true,
                    stroke: false,
                    fill_mode: None,
                }],
                text_area: None,
            };
            let Some((x0, y0, x1, y1)) = g.bounds() else {
                continue;
            };
            g.translate(-x0, -y0);
            let (mw, mh) = (x1 - x0, y1 - y0);
            let (cust, _) = normalize(&g, mw, mh);
            out.push(plain_shape(
                x0,
                y0,
                mw,
                mh,
                cust,
                Some(Fill::Solid {
                    color: marker.color.clone(),
                }),
                self.id(),
            ));
        }
        out
    }

    fn text(
        &mut self,
        node: Nd<'a>,
        ctx: &ObjectContext<'a>,
        styles: &ObjectStyles<'a>,
        graphic: &Props<'a>,
        background: Option<String>,
    ) -> Result<Option<TextBody>, String> {
        let container = node
            .children()
            .find(|c| xml::is(*c, NS_DRAW, "text-box"))
            .unwrap_or(node);
        let env = TextEnv {
            sheet: self.sheet,
            scope: ctx.scope,
            base: ctx.base,
            object: styles.clone(),
            slide_number: self.index + 1,
            slide_count: self.count,
            footer: ctx.footer.clone(),
            header: ctx.header.clone(),
            date_time: ctx.date_time.clone(),
            background,
            pages: self.page_names,
            payload: &text::TextPayload::new(self.text_work),
        };
        text::text_body(container, &env, self.media, graphic)
    }

    fn frame(
        &mut self,
        node: Nd<'a>,
        ctx: &ObjectContext<'a>,
        out: &mut Vec<SlideElement>,
    ) -> Result<(), String> {
        let styles = self.object_styles(node, ctx)?;
        let graphic = self.graphic_props(&styles)?;
        let geometry_node = self.placement_node(node, ctx)?;
        let (w, h) = Self::size(geometry_node);
        let placed = Self::placement(geometry_node, w, h);
        let children: Vec<Nd> = node.children().filter(|n| n.is_element()).collect();
        let has = |name: &str| children.iter().any(|c| c.tag_name().name() == name);
        if let Some(table) = children.iter().find(|c| xml::is(**c, NS_TABLE, "table")) {
            return crate::odp_table::table(
                self.sheet,
                self.media,
                *table,
                ctx.scope,
                ctx.base,
                &placed,
                &styles,
                self.index,
                self.count,
                self.page_names,
                self.text_work,
                out,
            );
        }
        if has("text-box") {
            let mut paint_ctx = PaintContext {
                sheet: self.sheet,
                media: self.media,
                base: ctx.base,
                slide_background: ctx.background.as_ref(),
            };
            let fill = paint::fill(&graphic, &mut paint_ctx, (placed.w, placed.h));
            let line = paint::line(&graphic, self.sheet);
            let background = solid_color(&fill).or_else(|| ctx.background_color.clone());
            let text_body = self.text(node, ctx, &styles, &graphic, background)?;
            let mut element = plain_shape(
                placed.x,
                placed.y,
                placed.w,
                placed.h,
                rect_path(),
                fill,
                self.id(),
            );
            if let SlideElement::Shape(s) = &mut element {
                s.rotation = placed.rot;
                s.flip_v = placed.flip_v;
                s.stroke = line.stroke;
                s.text_body = text_body;
                s.shadow = paint::shadow(&graphic);
                s.name = node.attribute((NS_DRAW, "name")).map(str::to_string);
                s.placeholder_type = placeholder_type(node);
                s.hyperlink = ctx.link.clone();
            }
            out.push(element);
            return Ok(());
        }
        // Charts (§10.4.2 draw:object holding an office:chart document) are
        // read into the shared chart model with their own data table
        // (`table:table` local to the chart). An object that is not a chart,
        // or a chart the reader cannot read, falls back to its replacement
        // picture below, as LibreOffice shows it.
        for child in &children {
            if !xml::is(*child, NS_DRAW, "object") {
                continue;
            }
            let Some(href) = child.attribute((NS_XLINK, "href")) else {
                continue;
            };
            let Some(crate::iri::Target::Package { path, .. }) =
                crate::iri::resolve(ctx.base, href)
            else {
                continue;
            };
            // Keyed by the object directory as the chart reader resolves it,
            // so spellings of one directory share one read.
            let path = path
                .trim_start_matches("./")
                .trim_end_matches('/')
                .to_string();
            let cached = match self.charts.get(&path) {
                Some(found) => found.clone(),
                None => {
                    let read =
                        crate::odf_chart::read_chart_object(self.media.package(), &path, None)
                            .ok()
                            .flatten();
                    // The cached model is retained for the whole read.
                    if let Some(model) = &read {
                        self.media.charge(chart_bytes(model)?)?;
                    }
                    self.charts.insert(path.clone(), read.clone());
                    read
                }
            };
            if let Some(chart) = cached {
                out.push(SlideElement::Chart(pptx_model::ChartElement {
                    id: Some(self.id()),
                    x: placed.x.round() as i64,
                    y: placed.y.round() as i64,
                    width: placed.w.round() as i64,
                    height: placed.h.round() as i64,
                    rotation: placed.rot,
                    flip_h: false,
                    flip_v: placed.flip_v,
                    chart,
                }));
                return Ok(());
            }
        }
        // Images, objects (charts, OLE) and plugins: the first renderable
        // image among the frame's alternatives (§10.4.2: the first child is
        // the preferred representation; later draw:image children are
        // replacements).
        let base = ctx.base;
        let mut chosen = None;
        for child in &children {
            if xml::is(*child, NS_DRAW, "image") {
                if let Some(href) = child.attribute((NS_XLINK, "href")) {
                    if let Some(found) = self.media.admit(base, href) {
                        chosen = Some(found);
                        break;
                    }
                }
            }
        }
        let Some((path, resource)) = chosen else {
            return Ok(());
        };
        let mirror = graphic.get(NS_STYLE, "mirror").unwrap_or("none");
        let mut blip_effects = Vec::new();
        let lum = graphic
            .get(NS_DRAW, "luminance")
            .and_then(percent)
            .unwrap_or(0.0);
        let contrast = graphic
            .get(NS_DRAW, "contrast")
            .and_then(percent)
            .unwrap_or(0.0);
        match graphic.get(NS_DRAW, "color-mode") {
            Some("greyscale") => blip_effects.push(BlipEffect::Grayscale),
            Some("mono") => blip_effects.push(BlipEffect::BiLevel { thresh: 0.5 }),
            // LibreOffice's watermark draw mode: luminance +50%, contrast −70%.
            Some("watermark") => blip_effects.push(BlipEffect::Luminance {
                bright: 0.5,
                contrast: -0.7,
            }),
            _ => {}
        }
        if lum != 0.0 || contrast != 0.0 {
            blip_effects.push(BlipEffect::Luminance {
                bright: lum,
                contrast,
            });
        }
        let opacity = graphic
            .get(NS_DRAW, "image-opacity")
            .and_then(percent)
            .unwrap_or(1.0);
        if opacity < 1.0 {
            blip_effects.push(BlipEffect::AlphaModFix {
                amt: opacity.max(0.0),
            });
        }
        let src_rect = graphic
            .get(NS_FO, "clip")
            .and_then(|clip| clip_rect(clip, &resource));
        let line = paint::line(&graphic, self.sheet);
        let picture = PictureElement {
            id: Some(self.id()),
            x: placed.x.round() as i64,
            y: placed.y.round() as i64,
            width: placed.w.round() as i64,
            height: placed.h.round() as i64,
            rotation: placed.rot,
            flip_h: mirror.contains("horizontal"),
            flip_v: mirror.contains("vertical") ^ placed.flip_v,
            image_path: path.clone(),
            mime_type: resource.mime.to_string(),
            svg_image_path: (resource.mime == "image/svg+xml").then(|| path.clone()),
            intrinsic_width_px: resource.pixels.map(|p| p.0),
            intrinsic_height_px: resource.pixels.map(|p| p.1),
            stroke: line.stroke,
            fill: None,
            prst_geom: None,
            prst_adjust: None,
            src_rect,
            alpha: None,
            duotone: None,
            blip_effects,
            cust_geom: None,
            shadow: paint::shadow(&graphic),
            inner_shadow: None,
            glow: paint::glow(&graphic),
            soft_edge: paint::soft_edge(&graphic),
            reflection: None,
            scene3d: None,
            sp3d: None,
        };
        out.push(SlideElement::Picture(picture));
        Ok(())
    }
}

/// LibreOffice's fontwork shape types and the DrawingML text warps they
/// correspond to (the pairs LibreOffice's OOXML filter converts between).
fn fontwork_warp(kind: &str) -> &'static str {
    match kind {
        "fontwork-stop" => "textStop",
        "fontwork-triangle-up" => "textTriangle",
        "fontwork-triangle-down" => "textTriangleInverted",
        "fontwork-chevron-up" => "textChevron",
        "fontwork-chevron-down" => "textChevronInverted",
        "fontwork-ring-inside" => "textRingInside",
        "fontwork-ring-outside" => "textRingOutside",
        "fontwork-arch-up-curve" => "textArchUp",
        "fontwork-arch-down-curve" => "textArchDown",
        "fontwork-circle-curve" => "textCircle",
        "fontwork-open-circle-curve" => "textButton",
        "fontwork-arch-up-pour" => "textArchUpPour",
        "fontwork-arch-down-pour" => "textArchDownPour",
        "fontwork-circle-pour" => "textCirclePour",
        "fontwork-open-circle-pour" => "textButtonPour",
        "fontwork-curve-up" => "textCurveUp",
        "fontwork-curve-down" => "textCurveDown",
        "fontwork-fade-up-and-right" => "textCascadeUp",
        "fontwork-fade-up-and-left" => "textCascadeDown",
        "fontwork-wave" => "textWave1",
        "mso-spt159" => "textWave2",
        "mso-spt160" => "textDoubleWave1",
        "fontwork-inflate" => "textInflate",
        "mso-spt161" => "textDeflate",
        "mso-spt162" => "textInflateBottom",
        "mso-spt163" => "textDeflateBottom",
        "mso-spt164" => "textInflateTop",
        "mso-spt165" => "textDeflateTop",
        "mso-spt166" => "textDeflateInflate",
        "mso-spt167" => "textDeflateInflateDeflate",
        "fontwork-fade-right" => "textFadeRight",
        "fontwork-fade-left" => "textFadeLeft",
        "fontwork-fade-up" => "textFadeUp",
        "fontwork-fade-down" => "textFadeDown",
        "fontwork-slant-up" => "textSlantUp",
        "fontwork-slant-down" => "textSlantDown",
        "mso-spt174" => "textCanUp",
        "mso-spt175" => "textCanDown",
        _ => "textPlain",
    }
}

fn clone_ctx<'a>(ctx: &ObjectContext<'a>) -> ObjectContext<'a> {
    ObjectContext {
        scope: ctx.scope,
        origin: ctx.origin,
        base: ctx.base,
        master_name: ctx.master_name,
        background: ctx.background.clone(),
        background_color: ctx.background_color.clone(),
        footer: ctx.footer.clone(),
        header: ctx.header.clone(),
        date_time: ctx.date_time.clone(),
        link: ctx.link.clone(),
        master: ctx.master,
    }
}

/// The shape a drawing object is: the one a hyperlink (`draw:a`) wraps,
/// else the object itself.
fn wrapped_shape(node: Nd<'_>) -> Nd<'_> {
    if xml::is(node, NS_DRAW, "a") {
        node.children().find(|c| c.is_element()).unwrap_or(node)
    } else {
        node
    }
}

/// Drawing objects in stacking order (§19.231 `draw:z-index`, ascending
/// from back to front; document order where it is absent or equal). An
/// object without a z-index stays right after the object before it.
fn stacking_order<'a>(nodes: Vec<Nd<'a>>) -> Vec<Nd<'a>> {
    // A hyperlink (`draw:a`) stacks as the shape it wraps.
    let z_index = |node: Nd<'a>| -> Option<u64> {
        wrapped_shape(node)
            .attribute((NS_DRAW, "z-index"))
            .and_then(|v| v.trim().parse::<u64>().ok())
    };
    let mut previous = 0u64;
    let mut keyed: Vec<(u64, usize, Nd<'a>)> = nodes
        .into_iter()
        .enumerate()
        .map(|(i, node)| {
            let key = z_index(node).unwrap_or(previous);
            previous = key;
            (key, i, node)
        })
        .collect();
    keyed.sort_by_key(|(key, i, _)| (*key, *i));
    keyed.into_iter().map(|(_, _, node)| node).collect()
}

/// In-memory bytes of one slide element: its fixed size, its serialized
/// size (the strings and lists it owns) and the fixed size of its repeated
/// records (paragraphs and runs, table rows and cells, path commands, chart
/// series, values and categories), which the serialized form understates.
fn model_bytes(element: &SlideElement) -> Result<usize, String> {
    use std::mem::size_of;
    let mut bytes = text::record_bytes(element)?;
    let mut add = |count: usize, size: usize| {
        bytes = bytes.saturating_add(count.saturating_mul(size));
    };
    let body = |body: &TextBody, add: &mut dyn FnMut(usize, usize)| {
        add(body.paragraphs.len(), size_of::<pptx_model::Paragraph>());
        for p in &body.paragraphs {
            add(p.runs.len(), size_of::<pptx_model::TextRun>());
        }
    };
    match element {
        SlideElement::Shape(shape) => {
            if let Some(b) = &shape.text_body {
                body(b, &mut add);
            }
            for path in shape.cust_geom.iter().flatten() {
                add(path.len(), size_of::<PathCmd>());
            }
        }
        SlideElement::Table(table) => {
            add(table.rows.len(), size_of::<pptx_model::TableRow>());
            for row in &table.rows {
                add(row.cells.len(), size_of::<pptx_model::TableCell>());
                for cell in &row.cells {
                    if let Some(b) = &cell.text_body {
                        body(b, &mut add);
                    }
                }
            }
        }
        SlideElement::Chart(chart) => {
            add(chart_records(&chart.chart), 1);
        }
        _ => {}
    }
    Ok(bytes)
}

/// Fixed size of a chart model's repeated records (series, values,
/// categories, trendlines), which its serialized form understates.
fn chart_records(model: &pptx_model::ChartModel) -> usize {
    use std::mem::size_of;
    let mut bytes = model.categories.len().saturating_mul(size_of::<String>());
    for series in &model.series {
        bytes = bytes
            .saturating_add(size_of::<ooxml_common::chart::ChartSeries>())
            .saturating_add(series.values.len().saturating_mul(size_of::<Option<f64>>()))
            .saturating_add(
                series
                    .categories
                    .as_ref()
                    .map_or(0, Vec::len)
                    .saturating_mul(size_of::<String>()),
            )
            .saturating_add(
                series
                    .trend_lines
                    .as_ref()
                    .map_or(0, Vec::len)
                    .saturating_mul(size_of::<ooxml_common::chart::ChartTrendline>()),
            );
        // The per-point vectors the reader fills (colours, bubble sizes,
        // hidden flags).
        let len = |v: Option<usize>, size: usize| v.unwrap_or(0).saturating_mul(size);
        bytes = bytes
            .saturating_add(len(
                series.data_point_colors.as_ref().map(Vec::len),
                size_of::<Option<String>>(),
            ))
            .saturating_add(len(
                series.data_label_colors.as_ref().map(Vec::len),
                size_of::<Option<String>>(),
            ))
            .saturating_add(len(
                series.bubble_sizes.as_ref().map(Vec::len),
                size_of::<Option<f64>>(),
            ))
            .saturating_add(len(
                series.source_hidden.as_ref().map(Vec::len),
                size_of::<bool>(),
            ))
            .saturating_add(len(
                series.cat_format_codes.as_ref().map(Vec::len),
                size_of::<Option<String>>(),
            ));
    }
    bytes = bytes.saturating_add(
        model
            .category_source_hidden
            .as_ref()
            .map_or(0, Vec::len)
            .saturating_mul(size_of::<bool>()),
    );
    bytes
}

/// In-memory bytes of a chart model (serialized size, fixed size and
/// repeated records).
fn chart_bytes(model: &pptx_model::ChartModel) -> Result<usize, String> {
    Ok(text::record_bytes(model)?.saturating_add(chart_records(model)))
}

fn placeholder_type(node: Nd<'_>) -> Option<String> {
    let class = node.attribute((NS_PRESENTATION, "class"))?;
    Some(
        match class {
            "title" => "title",
            "subtitle" => "subTitle",
            "outline" => "body",
            "footer" => "ftr",
            "header" => "hdr",
            "date-time" => "dt",
            "page-number" => "sldNum",
            "graphic" => "pic",
            "chart" => "chart",
            "table" => "tbl",
            "object" => "obj",
            "notes" => "body",
            _ => return None,
        }
        .to_string(),
    )
}

/// `fo:clip` (§20.179) "rect(top, right, bottom, left)" in the image's own
/// size, as source-rectangle fractions.
fn clip_rect(clip: &str, resource: &Resource) -> Option<SrcRect> {
    let inner = clip.trim().strip_prefix("rect(")?.strip_suffix(')')?;
    // Exactly four values: a fifth is rejected before it is read.
    let mut v = [0.0f64; 4];
    let mut n = 0;
    for s in inner
        .split(|c: char| c == ',' || c.is_whitespace())
        .filter(|s| !s.is_empty())
    {
        *v.get_mut(n)? = if s == "auto" { 0.0 } else { length_emu(s)? };
        n += 1;
    }
    if n != 4 {
        return None;
    }
    let (nw, nh) = resource.natural_emu()?;
    if nw <= 0.0 || nh <= 0.0 || v.iter().all(|x| *x == 0.0) {
        return None;
    }
    Some(SrcRect {
        t: v[0] / nh,
        r: v[1] / nw,
        b: v[2] / nh,
        l: v[3] / nw,
    })
}

fn solid_color(fill: &Option<Fill>) -> Option<String> {
    match fill {
        Some(Fill::Solid { color }) => Some(color[..6.min(color.len())].to_string()),
        _ => None,
    }
}

fn single(segs: Vec<Seg>, fill: bool) -> Geometry {
    Geometry {
        paths: vec![GeoPath {
            segs,
            fill,
            stroke: true,
            fill_mode: None,
        }],
        text_area: None,
    }
}

fn view_map(vb: Option<(f64, f64, f64, f64)>, w: f64, h: f64, x: f64, y: f64) -> (f64, f64) {
    match vb {
        Some((vx, vy, vw, vh)) if vw != 0.0 && vh != 0.0 => ((x - vx) * w / vw, (y - vy) * h / vh),
        // A path without a view box is in 1/100 mm.
        _ => (x * 360.0, y * 360.0),
    }
}

fn rect_geometry(w: f64, h: f64, radius: f64) -> Geometry {
    let r = radius.min(w / 2.0).min(h / 2.0).max(0.0);
    if r <= 0.0 {
        return single(
            vec![
                Seg::Move(0.0, 0.0),
                Seg::Line(w, 0.0),
                Seg::Line(w, h),
                Seg::Line(0.0, h),
                Seg::Close,
            ],
            true,
        );
    }
    let mut segs = vec![Seg::Move(r, 0.0), Seg::Line(w - r, 0.0)];
    let half = std::f64::consts::FRAC_PI_2;
    geo::arc_segments(w - r, r, r, r, -half, half, &mut segs);
    segs.push(Seg::Line(w, h - r));
    geo::arc_segments(w - r, h - r, r, r, 0.0, half, &mut segs);
    segs.push(Seg::Line(r, h));
    geo::arc_segments(r, h - r, r, r, half, half, &mut segs);
    segs.push(Seg::Line(0.0, r));
    geo::arc_segments(r, r, r, r, 2.0 * half, half, &mut segs);
    segs.push(Seg::Close);
    single(segs, true)
}

/// `draw:ellipse` / `draw:circle` with `draw:kind` full / section / cut / arc
/// and `draw:start-angle` / `draw:end-angle` (counter-clockwise degrees).
fn ellipse_geometry(node: Nd<'_>, w: f64, h: f64) -> Geometry {
    let (cx, cy, rx, ry) = (w / 2.0, h / 2.0, w / 2.0, h / 2.0);
    let kind = node.attribute((NS_DRAW, "kind")).unwrap_or("full");
    let angle = |name: &str, default: f64| {
        node.attribute((NS_DRAW, name))
            .and_then(crate::odp_units::angle_degrees)
            .unwrap_or(default)
    };
    let mut segs = Vec::new();
    if kind == "full" {
        segs.push(Seg::Move(cx + rx, cy));
        geo::arc_segments(cx, cy, rx, ry, 0.0, std::f64::consts::TAU, &mut segs);
        segs.push(Seg::Close);
        return single(segs, true);
    }
    let (a0, a1) = (angle("start-angle", 0.0), angle("end-angle", 360.0));
    let mut sweep = a1 - a0;
    if sweep <= 0.0 {
        sweep += 360.0;
    }
    let start = -a0.to_radians();
    let start_point = (cx + rx * start.cos(), cy + ry * start.sin());
    if kind == "section" {
        segs.push(Seg::Move(cx, cy));
        segs.push(Seg::Line(start_point.0, start_point.1));
    } else {
        segs.push(Seg::Move(start_point.0, start_point.1));
    }
    geo::arc_segments(cx, cy, rx, ry, start, -sweep.to_radians(), &mut segs);
    if kind != "arc" {
        segs.push(Seg::Close);
    }
    single(segs, kind != "arc")
}

fn regular_polygon(node: Nd<'_>, w: f64, h: f64) -> Geometry {
    let corners = node
        .attribute((NS_DRAW, "corners"))
        .and_then(|c| c.parse::<usize>().ok())
        .unwrap_or(3)
        .clamp(3, 1000);
    let concave = node.attribute((NS_DRAW, "concave")) == Some("true");
    let sharpness = node
        .attribute((NS_DRAW, "sharpness"))
        .and_then(percent)
        .unwrap_or(0.0);
    let (cx, cy, rx, ry) = (w / 2.0, h / 2.0, w / 2.0, h / 2.0);
    let mut segs = Vec::new();
    let n = if concave { corners * 2 } else { corners };
    for i in 0..n {
        let a = -std::f64::consts::FRAC_PI_2 + std::f64::consts::TAU * i as f64 / n as f64;
        let r = if concave && i % 2 == 1 {
            1.0 - sharpness
        } else {
            1.0
        };
        let p = (cx + rx * r * a.cos(), cy + ry * r * a.sin());
        segs.push(if i == 0 {
            Seg::Move(p.0, p.1)
        } else {
            Seg::Line(p.0, p.1)
        });
    }
    segs.push(Seg::Close);
    single(segs, true)
}

/// The range of the open subpath at the start or end of a painted path.
fn subpath_range(segs: &[Seg], at_start: bool) -> (usize, usize) {
    let moves: Vec<usize> = segs
        .iter()
        .enumerate()
        .filter(|(_, s)| matches!(s, Seg::Move(..)))
        .map(|(i, _)| i)
        .collect();
    if at_start {
        let start = moves.first().copied().unwrap_or(0);
        let end = moves.get(1).copied().unwrap_or(segs.len());
        (start, end)
    } else {
        let start = moves.last().copied().unwrap_or(0);
        (start, segs.len())
    }
}

fn rect_path() -> Vec<Vec<PathCmd>> {
    vec![vec![
        PathCmd::MoveTo { x: 0.0, y: 0.0 },
        PathCmd::LineTo { x: 1.0, y: 0.0 },
        PathCmd::LineTo { x: 1.0, y: 1.0 },
        PathCmd::LineTo { x: 0.0, y: 1.0 },
        PathCmd::Close,
    ]]
}

/// Normalize shape-local geometry to the model's [0, 1] box coordinates,
/// with per-path paint when a path is unfilled, unstroked or shaded.
fn normalize(geometry: &Geometry, w: f64, h: f64) -> (Vec<Vec<PathCmd>>, Option<Vec<PathPaint>>) {
    let nx = |x: f64| if w > 0.0 { x / w } else { 0.0 };
    let ny = |y: f64| if h > 0.0 { y / h } else { 0.0 };
    let mut paths = Vec::with_capacity(geometry.paths.len());
    let mut paint = Vec::with_capacity(geometry.paths.len());
    for path in &geometry.paths {
        let cmds = path
            .segs
            .iter()
            .map(|s| match *s {
                Seg::Move(x, y) => PathCmd::MoveTo { x: nx(x), y: ny(y) },
                Seg::Line(x, y) => PathCmd::LineTo { x: nx(x), y: ny(y) },
                Seg::Cubic(a, b, c, d, x, y) => PathCmd::CubicBezTo {
                    x1: nx(a),
                    y1: ny(b),
                    x2: nx(c),
                    y2: ny(d),
                    x: nx(x),
                    y: ny(y),
                },
                Seg::Quad(a, b, x, y) => PathCmd::QuadBezTo {
                    x1: nx(a),
                    y1: ny(b),
                    x: nx(x),
                    y: ny(y),
                },
                Seg::Close => PathCmd::Close,
            })
            .collect();
        paths.push(cmds);
        let fill = if !path.fill {
            Some("none".to_string())
        } else {
            path.fill_mode.map(str::to_string)
        };
        // LibreOffice fills a path's subpaths alternately (even-odd, the
        // VCL poly-polygon rule): a subpath inside another is a hole whatever
        // its winding, and self-intersections leave holes.
        paint.push(PathPaint {
            fill,
            stroke: path.stroke,
            fill_rule: Some("evenodd".to_string()),
        });
    }
    (paths, Some(paint))
}

fn plain_shape(
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    cust: Vec<Vec<PathCmd>>,
    fill: Option<Fill>,
    id: String,
) -> SlideElement {
    SlideElement::Shape(ShapeElement {
        math_fallback: None,
        x: x.round() as i64,
        y: y.round() as i64,
        width: w.round() as i64,
        height: h.round() as i64,
        rotation: 0.0,
        flip_h: false,
        flip_v: false,
        geometry: "custGeom".into(),
        fill,
        stroke: None,
        text_body: None,
        default_text_color: None,
        cust_geom: Some(cust),
        cust_geom_paint: None,
        adj: None,
        adj2: None,
        adj3: None,
        adj4: None,
        adj5: None,
        adj6: None,
        adj7: None,
        adj8: None,
        shadow: None,
        inner_shadow: None,
        glow: None,
        soft_edge: None,
        reflection: None,
        id: Some(id),
        name: None,
        hyperlink: None,
        hyperlink_action: None,
        placeholder_type: None,
        placeholder_idx: None,
        text_rect: None,
        scene3d: None,
        sp3d: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::tests::{manifest, package_bytes};
    use crate::package::PackageLimits;
    use ooxml_common::resource::OoxmlFormat;
    use pptx_model::Bullet;

    const NS: &str = r##"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0""##;
    const ODP: &str = "application/vnd.oasis.opendocument.presentation";

    /// A deck: `styles` (office:styles children), `master` (children of the
    /// "Default" master page), `automatic` (content automatic styles) and
    /// the `pages` (draw:page elements).
    fn deck(styles: &str, master: &str, automatic: &str, pages: &str) -> Result<OdpDeck, String> {
        let styles_xml = format!(
            r##"<office:document-styles {NS} office:version="1.3"><office:styles>{styles}</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="28cm" fo:page-height="21cm"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1">{master}</style:master-page></office:master-styles></office:document-styles>"##
        );
        let content = format!(
            r##"<office:document-content {NS} office:version="1.3"><office:automatic-styles>{automatic}</office:automatic-styles><office:body><office:presentation>{pages}</office:presentation></office:body></office:document-content>"##
        );
        let files = [
            ("META-INF/manifest.xml", manifest(ODP, "")),
            ("styles.xml", styles_xml),
            ("content.xml", content),
        ];
        let files: Vec<(&str, &[u8])> = files
            .iter()
            .map(|(name, body)| (*name, body.as_bytes()))
            .collect();
        let package = OdfPackage::open(
            package_bytes(ODP, &files),
            OoxmlFormat::Pptx,
            PackageLimits::default(),
            &MEDIA_TYPES,
        )?;
        read(&package)
    }

    /// The model of slide `index` (its serialized JSON is checked to be
    /// the bytes the deck keeps).
    fn model(deck: &OdpDeck, index: usize) -> Slide {
        let slide = deck.models[index].clone();
        assert_eq!(serde_json::to_vec(&slide).unwrap(), deck.slides[index].json);
        slide
    }

    fn page(body: &str) -> String {
        format!(r#"<draw:page draw:name="p1" draw:master-page-name="Default">{body}</draw:page>"#)
    }

    fn shapes(slide: &Slide) -> Vec<&ShapeElement> {
        slide
            .elements
            .iter()
            .filter_map(|e| match e {
                SlideElement::Shape(s) => Some(s),
                _ => None,
            })
            .collect()
    }

    fn labels(body: &TextBody) -> Vec<String> {
        body.paragraphs
            .iter()
            .map(|p| match &p.bullet {
                Bullet::Char { ch, .. } => ch.clone(),
                _ => String::new(),
            })
            .collect()
    }

    #[test]
    fn list_numbering_follows_list_identity() {
        let styles = r#"<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>"#;
        let item = |t: &str| format!("<text:list-item><text:p>{t}</text:p></text:list-item>");
        let body = format!(
            r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1" xml:id="a">{}{}</text:list><text:list text:style-name="L1" text:continue-numbering="false">{}</text:list><text:list text:style-name="L1" text:continue-numbering="true">{}</text:list><text:list text:style-name="L1" text:continue-list="a">{}</text:list><text:list text:style-name="L1">{}</text:list></draw:text-box></draw:frame>"#,
            item("a"),
            item("b"),
            item("c"),
            item("d"),
            item("e"),
            item("f")
        );
        let deck = deck(styles, "", "", &page(&body)).unwrap();
        let slide0 = model(&deck, 0);
        let shape = shapes(&slide0)[0];
        assert_eq!(
            labels(shape.text_body.as_ref().unwrap()),
            ["1.", "2.", "1.", "2.", "3.", "4."]
        );
    }

    #[test]
    fn item_paragraphs_and_nested_continuations_keep_numbering() {
        let styles = r#"<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/><text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")"/></text:list-style>"#;
        let body = r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1"><text:list-item><text:p>A</text:p><text:p>A, continued</text:p><text:list xml:id="s"><text:list-item><text:p>a</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>B</text:p><text:list text:continue-list="s"><text:list-item><text:p>b</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>C</text:p><text:list><text:list-item><text:p>c</text:p></text:list-item></text:list></text:list-item></text:list></draw:text-box></draw:frame>"#;
        let deck = deck(styles, "", "", &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        let shape = shapes(&slide0)[0];
        assert_eq!(
            labels(shape.text_body.as_ref().unwrap()),
            ["1.", "", "a)", "2.", "b)", "3.", "a)"]
        );
    }

    #[test]
    fn nested_lists_continue_or_restart_as_they_declare() {
        let styles = r#"<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/><text:list-level-style-number text:level="2" style:num-format="1" style:num-suffix=")"/></text:list-style>"#;
        let item = |t: &str| format!("<text:list-item><text:p>{t}</text:p></text:list-item>");
        let body = format!(
            r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1"><text:list-item><text:p>P</text:p><text:list xml:id="a">{}{}</text:list><text:list><text:list-item text:start-value="10"><text:p>s</text:p></text:list-item></text:list><text:list text:continue-list="a">{}</text:list><text:list text:continue-numbering="false">{}</text:list></text:list-item></text:list></draw:text-box></draw:frame>"#,
            item("x"),
            item("y"),
            item("z"),
            item("r")
        );
        let deck = deck(styles, "", "", &page(&body)).unwrap();
        let slide0 = model(&deck, 0);
        let shape = shapes(&slide0)[0];
        assert_eq!(
            labels(shape.text_body.as_ref().unwrap()),
            ["1.", "1)", "2)", "10)", "3)", "1)"]
        );
    }

    #[test]
    fn a_list_continues_a_list_of_another_frame() {
        let styles = r#"<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>"#;
        let item = |t: &str| format!("<text:list-item><text:p>{t}</text:p></text:list-item>");
        let body = format!(
            r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1" xml:id="a">{}{}</text:list></draw:text-box></draw:frame><draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1" text:continue-list="a">{}</text:list></draw:text-box></draw:frame>"#,
            item("x"),
            item("y"),
            item("z")
        );
        let deck = deck(styles, "", "", &page(&body)).unwrap();
        let slide0 = model(&deck, 0);
        let second = shapes(&slide0)[1];
        assert_eq!(labels(second.text_body.as_ref().unwrap()), ["3."]);
    }

    #[test]
    fn continuation_compares_the_effective_numbering_styles() {
        let styles = r#"<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style><text:list-style style:name="L2"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix=")" text:start-value="10"/></text:list-style><style:style style:name="P2" style:family="paragraph" style:list-style-name="L2"/>"#;
        let body = r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1"><text:list-item><text:p>a</text:p></text:list-item></text:list><text:list text:continue-numbering="true"><text:list-item><text:p text:style-name="P2">b</text:p></text:list-item></text:list></draw:text-box></draw:frame>"#;
        let deck = deck(styles, "", "", &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        let shape = shapes(&slide0)[0];
        assert_eq!(labels(shape.text_body.as_ref().unwrap()), ["1.", "10)"]);
    }

    #[test]
    fn continuation_sees_the_object_list_style() {
        let styles = r#"<text:list-style style:name="A"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style><style:style style:name="PA" style:family="paragraph" style:list-style-name="A"/>"#;
        let body = r#"<draw:frame svg:width="10cm" svg:height="5cm" draw:text-style-name="PA"><draw:text-box><text:list text:style-name="A"><text:list-item><text:p>a</text:p></text:list-item></text:list><text:list text:continue-numbering="true"><text:list-item><text:p>b</text:p></text:list-item></text:list></draw:text-box></draw:frame>"#;
        let deck = deck(styles, "", "", &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        let shape = shapes(&slide0)[0];
        assert_eq!(labels(shape.text_body.as_ref().unwrap()), ["1.", "2."]);
    }

    #[test]
    fn speaker_notes_read_each_paragraph_once() {
        let body = r#"<presentation:notes><draw:frame presentation:class="notes"><draw:text-box><text:p>One<text:s text:c="2"/>two</text:p><text:p>Three<text:p>nested</text:p></text:p></draw:text-box></draw:frame></presentation:notes>"#;
        let deck = deck("", "", "", &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        assert_eq!(slide0.notes.as_deref(), Some("One  two\nThreenested"));
    }

    #[test]
    fn a_local_continuation_follows_the_named_list() {
        let styles = r#"<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>"#;
        let body = r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1" xml:id="a"><text:list-item><text:p>x</text:p></text:list-item><text:list-item><text:p>y</text:p></text:list-item></text:list><text:list text:style-name="L1" text:continue-numbering="true"><text:list-item text:start-value="10"><text:p>s</text:p></text:list-item></text:list><text:list text:style-name="L1" text:continue-list="a"><text:list-item><text:p>z</text:p></text:list-item></text:list></draw:text-box></draw:frame>"#;
        let deck = deck(styles, "", "", &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        let shape = shapes(&slide0)[0];
        assert_eq!(
            labels(shape.text_body.as_ref().unwrap()),
            ["1.", "2.", "10.", "3."]
        );
    }

    #[test]
    fn a_continued_list_counts_its_own_continuation() {
        let styles = r#"<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>"#;
        let item = |t: &str| format!("<text:list-item><text:p>{t}</text:p></text:list-item>");
        let two = format!("{}{}", item("x"), item("y"));
        let body = format!(
            r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1">{two}</text:list><text:list text:style-name="L1" xml:id="b" text:continue-numbering="true">{two}</text:list></draw:text-box></draw:frame><draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box><text:list text:style-name="L1" text:continue-list="b">{}</text:list></draw:text-box></draw:frame>"#,
            item("z")
        );
        let deck = deck(styles, "", "", &page(&body)).unwrap();
        let slide0 = model(&deck, 0);
        let s = shapes(&slide0);
        assert_eq!(
            labels(s[0].text_body.as_ref().unwrap()),
            ["1.", "2.", "3.", "4."]
        );
        assert_eq!(labels(s[1].text_body.as_ref().unwrap()), ["5."]);
    }

    #[test]
    fn a_hyperlink_stacks_as_the_shape_it_wraps() {
        let body = r#"<draw:a xlink:href="https://example.com/"><draw:rect draw:name="linked" draw:z-index="10" svg:width="1cm" svg:height="1cm"/></draw:a><draw:rect draw:name="plain" draw:z-index="1" svg:width="1cm" svg:height="1cm"/>"#;
        let deck = deck("", "", "", &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        let names: Vec<_> = shapes(&slide0)
            .iter()
            .map(|s| s.name.clone().unwrap())
            .collect();
        assert_eq!(names, ["plain", "linked"]);
    }

    #[test]
    fn huge_table_spans_are_clamped_to_the_grid() {
        let body = r#"<draw:frame svg:width="10cm" svg:height="5cm"><table:table><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell table:number-columns-spanned="4294967295" table:number-rows-spanned="4294967295"><text:p>x</text:p></table:table-cell><table:covered-table-cell/></table:table-row><table:table-row><table:covered-table-cell table:number-columns-repeated="2"/></table:table-row></table:table></draw:frame>"#;
        let deck = deck("", "", "", &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        let SlideElement::Table(table) = &slide0.elements[0] else {
            panic!("{:?}", slide0.elements);
        };
        let first = &table.rows[0].cells[0];
        assert_eq!((first.grid_span, first.row_span), (2, 2));
        assert!(table.rows[1].cells[1].v_merge);
    }

    #[test]
    fn table_cell_padding_sets_the_cell_insets() {
        let automatic = r#"<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:padding="0.1cm" fo:padding-left="0.3cm"/></style:style>"#;
        let body = r#"<draw:frame svg:width="10cm" svg:height="5cm"><table:table><table:table-column/><table:table-row><table:table-cell table:style-name="ce1"><text:p>x</text:p></table:table-cell></table:table-row></table:table></draw:frame>"#;
        let deck = deck("", "", automatic, &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        let SlideElement::Table(table) = &slide0.elements[0] else {
            panic!();
        };
        let body = table.rows[0].cells[0].text_body.as_ref().unwrap();
        assert_eq!(
            (body.l_ins, body.t_ins, body.r_ins),
            (108_000, 36_000, 36_000)
        );
    }

    #[test]
    fn untransformed_placeholders_take_the_master_placement() {
        let master = r#"<draw:frame presentation:class="title" svg:x="2cm" svg:y="1cm" svg:width="20cm" svg:height="3cm"><draw:text-box/></draw:frame>"#;
        let body = r#"<draw:frame presentation:class="title" presentation:user-transformed="false" svg:x="9cm" svg:y="9cm" svg:width="1cm" svg:height="1cm"><draw:text-box><text:p>T</text:p></draw:text-box></draw:frame><draw:frame presentation:class="title" presentation:user-transformed="true" svg:x="9cm" svg:y="9cm" svg:width="1cm" svg:height="1cm"><draw:text-box><text:p>U</text:p></draw:text-box></draw:frame>"#;
        let deck = deck("", master, "", &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        let s = shapes(&slide0);
        assert_eq!(
            (s[0].x, s[0].y, s[0].width, s[0].height),
            (720_000, 360_000, 7_200_000, 1_080_000)
        );
        assert_eq!((s[1].x, s[1].width), (3_240_000, 360_000));
    }

    #[test]
    fn z_index_orders_objects_back_to_front() {
        let frame = |name: &str, z: &str| {
            format!(r#"<draw:rect draw:name="{name}" {z} svg:width="1cm" svg:height="1cm"/>"#)
        };
        let body = [
            frame("c", r#"draw:z-index="2""#),
            frame("a", r#"draw:z-index="0""#),
            frame("a2", ""),
            frame("b", r#"draw:z-index="1""#),
        ]
        .concat();
        let deck = deck("", "", "", &page(&body)).unwrap();
        let slide0 = model(&deck, 0);
        let names: Vec<_> = shapes(&slide0)
            .iter()
            .map(|s| s.name.clone().unwrap())
            .collect();
        assert_eq!(names, ["a", "a2", "b", "c"]);
    }

    #[test]
    fn extrusion_depth_defaults_to_36pt() {
        let body = r#"<draw:custom-shape svg:width="2cm" svg:height="2cm"><draw:enhanced-geometry draw:extrusion="true" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 Z N"/></draw:custom-shape>"#;
        let deck = deck("", "", "", &page(body)).unwrap();
        let slide0 = model(&deck, 0);
        let sp3d = shapes(&slide0)[0].sp3d.as_ref().unwrap();
        assert_eq!(sp3d.extrusion_h, 457_200);
    }

    #[test]
    fn the_text_limit_covers_every_paragraph_of_an_object() {
        // Each paragraph alone is far below the limit; together they exceed it.
        let spaces = r#"<text:s text:c="1024"/>"#.repeat(2_000);
        let paragraphs = format!("<text:p>{spaces}</text:p>").repeat(20);
        let body = format!(
            r#"<draw:frame svg:width="10cm" svg:height="5cm"><draw:text-box>{paragraphs}</draw:text-box></draw:frame>"#
        );
        let error = deck("", "", "", &page(&body)).err().unwrap();
        assert!(error.contains("text limit"), "{error}");
    }

    #[test]
    fn a_master_hatch_is_generated_once_for_every_slide() {
        let styles = r##"<draw:hatch draw:name="h1" draw:style="double" draw:color="#000000" draw:distance="0.1cm" draw:rotation="450"/><style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="hatch" draw:fill-hatch-name="h1"/></style:style>"##;
        let master = r#"<draw:rect draw:style-name="gr1" svg:width="5cm" svg:height="5cm"/>"#;
        let pages = page("").repeat(5);
        let deck = deck(styles, master, "", &pages).unwrap();
        assert!(matches!(
            shapes(&model(&deck, 4))[0].fill,
            Some(Fill::Image { .. })
        ));
        assert_eq!(deck.slides.len(), 5);
        assert_eq!(deck.converted.len(), 1);
    }
}
