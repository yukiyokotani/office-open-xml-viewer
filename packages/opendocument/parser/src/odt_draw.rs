//! Frames, pictures, text boxes and shapes (ODF Part 3 §9, §10) and their
//! graphic properties (§17.15 et al., §20.x `draw:`/`style:` attributes),
//! mapped onto DOCX pictures and shapes.
//!
//! Anchoring (§19.27 `text:anchor-type`): `as-char` objects flow in the line
//! as DOCX inline objects; `char`, `paragraph`, `frame` and `page` anchors
//! become DOCX anchors whose frames follow `style:horizontal-rel` /
//! `style:vertical-rel` (§20.291, §20.388) and whose offsets or alignments
//! follow `style:horizontal-pos` / `style:vertical-pos` (§20.290, §20.387).
//! Page-anchored objects in the body are carried by the following paragraph
//! (DOCX anchors live in paragraphs); their position stays page-relative.
//! Text wrapping (§20.384 `style:wrap`) maps onto the DOCX wrap kinds.

use super::*;
use docx_model::{
    AnchorAcquisitionWire, AnchorAxisChoiceWire, AnchorAxisWire, AnchorBehaviorWire,
    AnchorEdgesWire, AnchorExtentWire, AnchorSimplePositionWire, AnchorValueStatusWire,
    AnchorWrapKindWire, AnchorWrapWire, GradientStop, ImageRun, LineEnd, PathCmd, ShapeFill,
    ShapeRun, SrcRect, TextBoxBlockWire,
};

/// Drawing definitions of `office:styles` (§16.40-§16.46): gradients,
/// hatches, fill images, markers and dash styles, by `draw:name`.
#[derive(Default)]
pub(super) struct Definitions<'a> {
    /// Each definition with its source size (a use reads it again).
    by_name: HashMap<(&'static str, String), (roxmltree::Node<'a, 'a>, usize)>,
}

impl<'a> Definitions<'a> {
    pub fn read(styles: Option<roxmltree::Node<'a, 'a>>) -> Result<Self, String> {
        let mut by_name = HashMap::new();
        if let Some(styles) = styles {
            for node in styles.kids().filter(|n| n.is_element()) {
                budget::tick(1)?;
                if node.tag_name().namespace() != Some(NS_DRAW) {
                    continue;
                }
                let kind: &'static str = match node.tag_name().name() {
                    "gradient" => "gradient",
                    "linear-gradient" => "gradient",
                    "radial-gradient" => "gradient",
                    "hatch" => "hatch",
                    "fill-image" => "fill-image",
                    "marker" => "marker",
                    "stroke-dash" => "stroke-dash",
                    "opacity" => "opacity",
                    _ => continue,
                };
                if let Some(name) = node.attribute((NS_DRAW, "name")) {
                    by_name
                        .entry((kind, name.to_string()))
                        .or_insert((node, source_bytes(node)?));
                }
            }
        }
        Ok(Self { by_name })
    }

    pub(super) fn get(
        &self,
        kind: &'static str,
        name: &str,
    ) -> Option<(roxmltree::Node<'a, 'a>, usize)> {
        self.by_name.get(&(kind, name.to_string())).copied()
    }
}

/// A resolved graphic style (§16.2 over `style:graphic-properties`).
struct Graphic<'a> {
    cascade: Cascade<'a>,
    /// Where the paragraphs of the drawing shape or frame continue their
    /// property search: the paragraph style its `draw:text-style-name`
    /// names (§19.227) and that style's parents, then the graphic style and
    /// its parents (§16.2), without the families' default styles.
    chain: Rc<[StyleRef<'a>]>,
}

impl<'a> Graphic<'a> {
    fn get(&self, ns: &str, name: &str) -> Option<&'a str> {
        self.cascade.get(ns, name)
    }

    fn length(&self, ns: &str, name: &str) -> Option<f64> {
        self.get(ns, name).and_then(|v| length_pt(v).ok())
    }
}

/// An inline picture run (no anchor), shared with picture list labels.
pub(super) fn inline_image(
    path: String,
    mime_type: &str,
    width: f64,
    height: f64,
    list_label: Option<&'static str>,
) -> ImageRun {
    ImageRun {
        inline_object_font_facts: None,
        inline_object_origin: None,
        inline_effect_extent: None,
        image_path: path,
        mime_type: mime_type.to_string(),
        svg_image_path: None,
        src_rect: None,
        width_pt: width,
        height_pt: height,
        rotation: 0.0,
        flip_h: false,
        flip_v: false,
        anchor: false,
        anchor_x_pt: 0.0,
        anchor_y_pt: 0.0,
        anchor_x_from_margin: false,
        anchor_y_from_para: false,
        color_replace_from: None,
        duotone: None,
        blip_effects: Vec::new(),
        alpha: None,
        wrap_mode: None,
        dist_top: 0.0,
        dist_bottom: 0.0,
        dist_left: 0.0,
        dist_right: 0.0,
        wrap_side: None,
        allow_overlap: false,
        anchor_x_align: None,
        anchor_y_align: None,
        anchor_x_relative_from: None,
        anchor_y_relative_from: None,
        anchor_acquisition: None,
        list_label,
    }
}

/// The anchor of a floating object: frames and choices in DOCX terms.
struct Placement {
    inline: bool,
    horizontal: (String, AnchorAxisChoiceWire),
    vertical: (String, AnchorAxisChoiceWire),
    wrap: AnchorWrapKindWire,
    side: Option<&'static str>,
    behind: bool,
    distances: [f64; 4],
    z_order: u32,
    /// LibreOffice `loext:allow-overlap` (default true).
    allow_overlap: bool,
    /// `draw:contour-polygon` (§10.4.12) in the 21600 wrap-polygon space.
    polygon: Option<docx_model::AnchorPolygonWire>,
}

impl Placement {
    /// The wrap kind drawn: a contour wrap without its polygon wraps around
    /// the object's box.
    fn effective_wrap(&self) -> AnchorWrapKindWire {
        match (self.wrap, &self.polygon) {
            (AnchorWrapKindWire::Tight, None) => AnchorWrapKindWire::Square,
            (kind, _) => kind,
        }
    }
}

/// `style:wrap` (§20.384) and `style:wrap-contour` (§20.385).
fn wrap(graphic: &Graphic<'_>) -> (AnchorWrapKindWire, Option<&'static str>, bool) {
    let behind = graphic.get(NS_STYLE, "run-through") == Some("background");
    let contour = graphic.get(NS_STYLE, "wrap-contour") == Some("true");
    let square = if contour {
        AnchorWrapKindWire::Tight
    } else {
        AnchorWrapKindWire::Square
    };
    match graphic.get(NS_STYLE, "wrap") {
        Some("none") => (AnchorWrapKindWire::TopAndBottom, None, false),
        Some("left") => (square, Some("left"), false),
        Some("right") => (square, Some("right"), false),
        Some("parallel") => (square, Some("bothSides"), false),
        Some("dynamic" | "biggest") => (square, Some("largest"), false),
        Some("run-through") => (AnchorWrapKindWire::None, None, behind),
        _ => (AnchorWrapKindWire::None, None, behind),
    }
}

impl<'a> Reader<'a> {
    fn graphic(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<Graphic<'a>, String> {
        let own: Rc<[StyleRef<'a>]> = match node
            .attribute((NS_DRAW, "style-name"))
            .and_then(|name| self.find_style(Family::Graphic, name))
        {
            Some(style) => self.style_chain(style)?,
            None => Rc::from(Vec::new()),
        };
        let mut chain = own.to_vec();
        if let Some(default) = self.sheet.default_style(Family::Graphic) {
            chain.push(default);
        }
        let text_chain = match node
            .attribute((NS_DRAW, "text-style-name"))
            .and_then(|name| self.find_style(Family::Paragraph, name))
        {
            Some(style) => {
                let mut text_chain = self.style_chain(style)?.to_vec();
                text_chain.extend(own.iter().copied());
                text_chain.into()
            }
            None => own,
        };
        Ok(Graphic {
            cascade: Cascade::build(chain, "graphic-properties")?,
            chain: text_chain,
        })
    }

    /// A drawing element met in content: a run for the paragraph that
    /// carries it, or `None` when it draws nothing.
    pub(super) fn drawing(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        _host: Option<(u32, &Inline<'a>)>,
    ) -> Result<Option<DocRun>, String> {
        // §19.758-§19.759: a page-anchored drawing with an explicit physical
        // page belongs to that page, wherever it is written (body, paragraph
        // or a master page's header or footer, as LibreOffice reads it): it
        // becomes a page object, not content of the story that holds it.
        // A group member's anchor is its group's (the group itself is the
        // page object, or not one).
        let page = if self.group_depth > 0 {
            None
        } else {
            self.anchor_page_number(node)?
        };
        if let Some(page) = page {
            // One page object per drawing element, however often the story
            // holding it is read (a master's header can be read for several
            // page styles).
            let identity = (node.document() as *const _ as usize, node.id());
            if !self.page_object_nodes.insert(identity) {
                return Ok(None);
            }
            // A group's further members join the pending anchors: they belong
            // to the page too.
            let before = self.pending_anchors.len();
            let run = self.drawing_run(node)?;
            let members = self.pending_anchors.split_off(before);
            let objects = self.page_objects.entry(page).or_default();
            objects.extend(run);
            objects.extend(members);
            return Ok(None);
        }
        self.drawing_run(node)
    }

    /// `text:anchor-page-number` (§19.758) of a page-anchored drawing
    /// (§19.759 `text:anchor-type="page"`), on the element or from its
    /// graphic style (§20.417-§20.418): its physical page, 1-based. Page
    /// numbers past the reader's page bound are refused (LibreOffice extends
    /// the document to the page, so the bound caps the pages a drawing adds).
    fn anchor_page_number(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<Option<u32>, String> {
        let name = node.tag_name().name();
        if node.tag_name().namespace() != Some(NS_DRAW) || name == "a" {
            return Ok(None);
        }
        let own = |attribute: &str| node.attribute((NS_TEXT, attribute));
        let (anchor, page) = if own("anchor-type").is_some() && own("anchor-page-number").is_some()
        {
            (own("anchor-type"), own("anchor-page-number"))
        } else {
            let graphic = self.graphic(node)?;
            (
                own("anchor-type").or_else(|| graphic.get(NS_TEXT, "anchor-type")),
                own("anchor-page-number").or_else(|| graphic.get(NS_TEXT, "anchor-page-number")),
            )
        };
        if anchor != Some("page") {
            return Ok(None);
        }
        Ok(
            xml::bounded_count(page, u64::from(MAX_ANCHOR_PAGE), "page anchor")?
                .filter(|page| *page >= 1)
                .map(|page| page as u32),
        )
    }

    fn drawing_run(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<Option<DocRun>, String> {
        self.enter()?;
        let result = match node.tag_name().name() {
            "frame" => self.frame(node),
            // A hyperlinked drawing (§9.7.1): its frame.
            "a" => match node.kids().find(|n| n.is_element()) {
                Some(child) => self.drawing(child, None),
                None => Ok(None),
            },
            "custom-shape" | "rect" | "ellipse" | "circle" | "line" | "polyline" | "polygon"
            | "path" | "connector" | "caption" | "measure" | "regular-polygon" => {
                self.shape(node, None)
            }
            "g" => self.group(node),
            _ => Ok(None),
        };
        self.leave();
        result
    }

    /// Anchoring and wrapping of a frame or shape.
    fn placement(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        graphic: &Graphic<'a>,
        offset: (f64, f64),
    ) -> Result<Placement, String> {
        // A group member has no anchor of its own: it takes its group's
        // (§10.3.15; the group is the anchored object).
        let anchor = self
            .group_anchor
            .or_else(|| node.attribute((NS_TEXT, "anchor-type")))
            .or_else(|| graphic.get(NS_TEXT, "anchor-type"))
            .unwrap_or("paragraph");
        let x = node
            .attribute((NS_SVG, "x"))
            .and_then(|v| length_pt(v).ok())
            .unwrap_or(0.0)
            + offset.0;
        let y = node
            .attribute((NS_SVG, "y"))
            .and_then(|v| length_pt(v).ok())
            .unwrap_or(0.0)
            + offset.1;
        let (wrap, side, behind) = wrap(graphic);
        let distances = [
            graphic.length(NS_FO, "margin-top").unwrap_or(0.0).max(0.0),
            graphic
                .length(NS_FO, "margin-right")
                .unwrap_or(0.0)
                .max(0.0),
            graphic
                .length(NS_FO, "margin-bottom")
                .unwrap_or(0.0)
                .max(0.0),
            graphic.length(NS_FO, "margin-left").unwrap_or(0.0).max(0.0),
        ];
        let z_order = node.attribute((NS_DRAW, "z-index"));
        // (DOCX relative heights above the text start at 251658240.)
        let z_order = xml::bounded_count(z_order, u64::from(u32::MAX) - 251_658_240, "z-index")?
            .map_or(0, |index| index as u32)
            + 251_658_240;
        let page_anchor = anchor == "page";
        let horizontal_rel = graphic
            .get(NS_STYLE, "horizontal-rel")
            .unwrap_or(if page_anchor { "page" } else { "paragraph" });
        let horizontal_frame = match horizontal_rel {
            "page" => "page",
            "page-content" => "margin",
            "page-start-margin" | "page-start-margin-area" => "leftMargin",
            "page-end-margin" | "page-end-margin-area" => "rightMargin",
            "char" => "character",
            _ => "column",
        };
        let horizontal = match graphic
            .get(NS_STYLE, "horizontal-pos")
            .unwrap_or("from-left")
        {
            "left" => AnchorAxisChoiceWire::Align {
                value: "left".into(),
            },
            "center" => AnchorAxisChoiceWire::Align {
                value: "center".into(),
            },
            "right" => AnchorAxisChoiceWire::Align {
                value: "right".into(),
            },
            "inside" => AnchorAxisChoiceWire::Align {
                value: "inside".into(),
            },
            "outside" => AnchorAxisChoiceWire::Align {
                value: "outside".into(),
            },
            _ => AnchorAxisChoiceWire::Offset { value_pt: x },
        };
        let vertical_rel = graphic
            .get(NS_STYLE, "vertical-rel")
            .unwrap_or(if page_anchor { "page" } else { "paragraph" });
        let vertical_frame = match vertical_rel {
            "page" => "page",
            "page-content" => "margin",
            "char" | "line" | "baseline" | "text" => "line",
            _ => "paragraph",
        };
        let vertical = match graphic.get(NS_STYLE, "vertical-pos").unwrap_or("from-top") {
            "top" => AnchorAxisChoiceWire::Align {
                value: "top".into(),
            },
            "middle" => AnchorAxisChoiceWire::Align {
                value: "center".into(),
            },
            "bottom" => AnchorAxisChoiceWire::Align {
                value: "bottom".into(),
            },
            _ => AnchorAxisChoiceWire::Offset { value_pt: y },
        };
        Ok(Placement {
            inline: anchor == "as-char",
            horizontal: (horizontal_frame.into(), horizontal),
            vertical: (vertical_frame.into(), vertical),
            wrap,
            side,
            behind,
            distances,
            z_order,
            allow_overlap: graphic.get(NS_LOEXT, "allow-overlap") != Some("false"),
            polygon: contour_polygon(node)?,
        })
    }

    fn acquisition(
        &mut self,
        placement: &Placement,
        width: f64,
        height: f64,
    ) -> AnchorAcquisitionWire {
        let valid = |value: f64| (Some(value), AnchorValueStatusWire::Valid);
        let axis = |(frame, choice): &(String, AnchorAxisChoiceWire)| AnchorAxisWire {
            relative_from: Some(frame.clone()),
            relative_from_status: AnchorValueStatusWire::Valid,
            choice: choice.clone(),
        };
        let [top, right, bottom, left] = placement.distances;
        let edges = AnchorEdgesWire {
            top_pt: valid(top).0,
            top_status: AnchorValueStatusWire::Valid,
            right_pt: valid(right).0,
            right_status: AnchorValueStatusWire::Valid,
            bottom_pt: valid(bottom).0,
            bottom_status: AnchorValueStatusWire::Valid,
            left_pt: valid(left).0,
            left_status: AnchorValueStatusWire::Valid,
        };
        AnchorAcquisitionWire {
            occurrence_id: self.occurrence("anchor"),
            simple_position: AnchorSimplePositionWire {
                enabled: Some(false),
                status: AnchorValueStatusWire::Valid,
                x_pt: Some(0.0),
                x_status: AnchorValueStatusWire::Valid,
                y_pt: Some(0.0),
                y_status: AnchorValueStatusWire::Valid,
            },
            horizontal: axis(&placement.horizontal),
            vertical: axis(&placement.vertical),
            extent: AnchorExtentWire {
                width_pt: Some(width),
                height_pt: Some(height),
                width_status: AnchorValueStatusWire::Valid,
                height_status: AnchorValueStatusWire::Valid,
            },
            anchor_distances: edges.clone(),
            wrap: AnchorWrapWire {
                // A contour wrap needs its polygon; without one it wraps
                // around the frame's box.
                kind: placement.effective_wrap(),
                polygon: placement
                    .polygon
                    .clone()
                    .filter(|_| placement.wrap == AnchorWrapKindWire::Tight),
                authored_kinds: vec![match placement.effective_wrap() {
                    AnchorWrapKindWire::TopAndBottom => "wrapTopAndBottom",
                    AnchorWrapKindWire::Square => "wrapSquare",
                    AnchorWrapKindWire::Tight => "wrapTight",
                    _ => "wrapNone",
                }
                .into()],
                side: placement.side.map(str::to_string),
                distances: edges,
                effect_extent: None,
            },
            behavior: AnchorBehaviorWire {
                behind_doc: Some(placement.behind),
                behind_doc_status: AnchorValueStatusWire::Valid,
                relative_height: Some(placement.z_order),
                relative_height_status: AnchorValueStatusWire::Valid,
                locked: Some(false),
                locked_status: AnchorValueStatusWire::Valid,
                allow_overlap: Some(placement.allow_overlap),
                allow_overlap_status: AnchorValueStatusWire::Valid,
                layout_in_cell: Some(true),
                layout_in_cell_status: AnchorValueStatusWire::Valid,
            },
            ..AnchorAcquisitionWire::default()
        }
    }

    /// A frame's size (§19.575 svg:width, §19.543 svg:height, §20.188
    /// fo:min-height). A percentage (§19.514 style:rel-width, §19.513
    /// style:rel-height) is the size; the lengths written beside it are the
    /// producer's value for consumers without relative sizes. LibreOffice
    /// takes the percentage of the text area holding the frame, or of the
    /// whole page with `loext:rel-width-rel="page"` (its reading of the
    /// basis is followed).
    fn frame_size(&self, node: roxmltree::Node<'a, 'a>, graphic: &Graphic<'a>) -> (f64, f64) {
        let get = |ns: &str, name: &str| node.attribute((ns, name)).and_then(|v| length_pt(v).ok());
        let rel = |name: &str| {
            node.attribute((NS_STYLE, name))
                .filter(|v| v.ends_with('%'))
                .and_then(|v| percent(v).ok())
                .filter(|share| *share > 0.0)
        };
        let page = |name: &str| node.attribute((NS_LOEXT, name)) == Some("page");
        let section = self.pages.page_size(&self.current_master);
        let width = rel("rel-width")
            .map(|r| {
                r * if page("rel-width-rel") {
                    section.0
                } else {
                    self.text_width
                }
            })
            .or_else(|| get(NS_SVG, "width"))
            .or_else(|| get(NS_FO, "min-width"))
            .or_else(|| graphic.length(NS_FO, "min-width"))
            .unwrap_or(0.0);
        let height = rel("rel-height")
            .map(|r| {
                r * if page("rel-height-rel") {
                    section.1
                } else {
                    self.text_height
                }
            })
            .or_else(|| get(NS_SVG, "height"))
            .or_else(|| get(NS_FO, "min-height"))
            .or_else(|| graphic.length(NS_FO, "min-height"))
            .unwrap_or(0.0);
        (width.max(0.0), height.max(0.0))
    }

    /// `draw:frame` (§10.4.2): a picture, a text box or an object's
    /// replacement picture.
    fn frame(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<Option<DocRun>, String> {
        let graphic = self.graphic(node)?;
        let (width, mut height) = self.frame_size(node, &graphic);
        let text_box = node.kids().find(|n| xml::is(*n, NS_DRAW, "text-box"));
        // §19.240: a text box's `fo:min-height` replaces the frame's height
        // (before the frame is positioned by its transform); zero is a
        // height like any other, and the box grows from it. A text box
        // with no height at all starts at one line.
        if text_box.is_some() {
            match text_box
                .and_then(|text_box| text_box.attribute((NS_FO, "min-height")))
                .and_then(|v| length_pt(v).ok())
                .filter(|minimum| *minimum >= 0.0)
            {
                Some(minimum) => height = minimum,
                None if height <= 0.0 => height = 12.0,
                None => {}
            }
        }
        let mut placement = self.placement(node, &graphic, (0.0, 0.0))?;
        // `draw:transform` (§19.228) positions the frame like a shape: the
        // transformed box's corner replaces the offsets of `svg:x` and
        // `svg:y`.
        let at = |name: &str| {
            node.attribute((NS_SVG, name))
                .and_then(|v| length_pt(v).ok())
                .unwrap_or(0.0)
        };
        let (rotation, translated) = transform_rotation(
            node.attribute((NS_DRAW, "transform")),
            (at("x"), at("y")),
            width,
            height,
        );
        if let Some((x, y)) = translated {
            if let AnchorAxisChoiceWire::Offset { value_pt } = &mut placement.horizontal.1 {
                *value_pt = x;
            }
            if let AnchorAxisChoiceWire::Offset { value_pt } = &mut placement.vertical.1 {
                *value_pt = y;
            }
        }
        if let Some(text_box) = text_box {
            return self.text_box(node, text_box, &graphic, placement, width, height, rotation);
        }
        // A formula object (MathML) draws as math: in the line when the
        // frame is anchored as a character, else as the content of the
        // anchored frame, which keeps its position, size and wrap.
        if let Some(math) = self.formula_object(node)? {
            if placement.inline {
                return Ok(Some(math));
            }
            self.count_paragraph()?;
            let mut shape = ShapeRun {
                width_pt: width,
                height_pt: height,
                rotation,
                preset_geometry: Some("rect".into()),
                text_box_content: vec![TextBoxBlockWire::Body(BodyElement::Paragraph(Box::new(
                    DocParagraph {
                        alignment: "left".into(),
                        runs: vec![math],
                        ..DocParagraph::default()
                    },
                )))],
                text_anchor: Some("t".into()),
                text_autofit: Some("none".into()),
                text_inset_l: 0.0,
                text_inset_t: 0.0,
                text_inset_r: 0.0,
                text_inset_b: 0.0,
                ..ShapeRun::default()
            };
            self.anchor_shape(&mut shape, &placement);
            return Ok(Some(DocRun::Shape(Box::new(shape))));
        }
        // The pictures of the frame: the first renderable one, with an SVG
        // original kept beside its raster replacement.
        let mut raster: Option<(String, &'static str)> = None;
        let mut svg: Option<String> = None;
        for image in node.kids().filter(|n| xml::is(*n, NS_DRAW, "image")).chain(
            node.kids()
                .filter(|n| xml::is(*n, NS_DRAW, "object") || xml::is(*n, NS_DRAW, "object-ole"))
                .flat_map(|object| object.kids().filter(|n| xml::is(*n, NS_DRAW, "image"))),
        ) {
            let resolved = match image.attribute((NS_XLINK, "href")) {
                Some(href) => self.package_image(href)?,
                None => match image
                    .kids()
                    .find(|n| xml::is(*n, NS_OFFICE, "binary-data"))
                    .and_then(|n| n.text())
                {
                    Some(data) => self.binary_image(data)?,
                    None => None,
                },
            };
            let Some((path, mime)) = resolved else {
                continue;
            };
            if mime == "image/svg+xml" {
                svg.get_or_insert(path);
            } else if raster.is_none() {
                raster = Some((path, mime));
            }
        }
        let (path, mime) = match (raster, svg.clone()) {
            (Some(raster), _) => raster,
            (None, Some(svg)) => (svg, "image/svg+xml"),
            (None, None) => return Ok(None),
        };
        // Padding and border inset the picture inside the frame.
        let padding = |side: &str| {
            graphic
                .length(NS_FO, &format!("padding-{side}"))
                .or_else(|| graphic.length(NS_FO, "padding"))
                .unwrap_or(0.0)
        };
        let image_width = (width - padding("left") - padding("right")).max(0.1);
        let image_height = (height - padding("top") - padding("bottom")).max(0.1);
        let mut image = inline_image(path.clone(), mime, image_width, image_height, None);
        image.svg_image_path = svg.filter(|svg| *svg != path);
        image.rotation = rotation;
        // §20.351 style:mirror.
        let mirror = graphic.get(NS_STYLE, "mirror").unwrap_or("none");
        image.flip_h = mirror.contains("horizontal");
        image.flip_v = mirror.contains("vertical");
        // §20.172 fo:clip: rect(top, right, bottom, left) of the picture's
        // intrinsic size.
        if let Some(clip) = graphic.get(NS_FO, "clip") {
            image.src_rect = self.clip(&path, clip);
        }
        // §20.162 draw:image-opacity, §20.137 draw:color-mode, luminance and
        // contrast (§20.150, §20.135).
        if let Some(opacity) = graphic
            .get(NS_DRAW, "image-opacity")
            .and_then(|v| percent(v).ok())
            .filter(|v| *v < 1.0)
        {
            image.alpha = Some(opacity.max(0.0));
        }
        match graphic.get(NS_DRAW, "color-mode") {
            Some("greyscale") => image.blip_effects.push(docx_model::BlipEffect::Grayscale),
            Some("mono") => image
                .blip_effects
                .push(docx_model::BlipEffect::BiLevel { thresh: 0.5 }),
            Some("watermark") => image.blip_effects.push(docx_model::BlipEffect::Luminance {
                bright: 0.7,
                contrast: -0.7,
            }),
            _ => {}
        }
        let bright = graphic
            .get(NS_DRAW, "luminance")
            .and_then(|v| percent(v).ok())
            .unwrap_or(0.0);
        let contrast = graphic
            .get(NS_DRAW, "contrast")
            .and_then(|v| percent(v).ok())
            .unwrap_or(0.0);
        if bright != 0.0 || contrast != 0.0 {
            image
                .blip_effects
                .push(docx_model::BlipEffect::Luminance { bright, contrast });
        }
        if placement.inline {
            return Ok(Some(DocRun::Image(Box::new(image))));
        }
        image.anchor = true;
        self.anchor_image(&mut image, &placement, width, height);
        Ok(Some(DocRun::Image(Box::new(image))))
    }

    /// `draw:object` linking a formula sub-document (Part 3 §10.4.6.2): its
    /// MathML content as DOCX math.
    fn formula_object(&mut self, frame: roxmltree::Node<'a, 'a>) -> Result<Option<DocRun>, String> {
        let Some(href) = frame
            .kids()
            .find(|n| xml::is(*n, NS_DRAW, "object"))
            .and_then(|object| object.attribute((NS_XLINK, "href")))
        else {
            return Ok(None);
        };
        let Some(iri::Target::Package { path, .. }) = iri::resolve("content.xml", href) else {
            return Ok(None);
        };
        let content = format!("{}/content.xml", path.trim_end_matches('/'));
        if !self.package.contains(&content) {
            return Ok(None);
        }
        // Each referencing frame owns its tree: the sub-document is read
        // and converted per frame, with steps for its size.
        let text = self.package.read_text(&content)?;
        rescan(text.len())?;
        let document = xml::parse(&text)?;
        let root = document.root_element();
        if root.tag_name().name() != "math" {
            return Ok(None);
        }
        let nodes = math::convert(root)?;
        if nodes.is_empty() {
            return Ok(None);
        }
        Ok(Some(DocRun::Math {
            nodes,
            display: false,
            font_size: 12.0,
            jc: None,
        }))
    }

    fn anchor_image(
        &mut self,
        image: &mut ImageRun,
        placement: &Placement,
        width: f64,
        height: f64,
    ) {
        let acquisition = self.acquisition(placement, width, height);
        let (x, x_align) = match &placement.horizontal.1 {
            AnchorAxisChoiceWire::Offset { value_pt } => (*value_pt, None),
            AnchorAxisChoiceWire::Align { value } => (0.0, Some(value.clone())),
            _ => (0.0, None),
        };
        let (y, y_align) = match &placement.vertical.1 {
            AnchorAxisChoiceWire::Offset { value_pt } => (*value_pt, None),
            AnchorAxisChoiceWire::Align { value } => (0.0, Some(value.clone())),
            _ => (0.0, None),
        };
        image.anchor_x_pt = x;
        image.anchor_y_pt = y;
        image.anchor_x_from_margin = placement.horizontal.0 != "page";
        image.anchor_y_from_para = placement.vertical.0 == "paragraph";
        image.anchor_x_align = x_align;
        image.anchor_y_align = y_align;
        image.anchor_x_relative_from = Some(placement.horizontal.0.clone());
        image.anchor_y_relative_from = Some(placement.vertical.0.clone());
        image.wrap_mode = Some(wrap_mode(placement.effective_wrap()).into());
        image.wrap_side = placement.side.map(str::to_string);
        image.dist_top = placement.distances[0];
        image.dist_right = placement.distances[1];
        image.dist_bottom = placement.distances[2];
        image.dist_left = placement.distances[3];
        image.allow_overlap = placement.allow_overlap;
        image.anchor_acquisition = Some(acquisition);
    }

    fn clip(&self, path: &str, clip: &str) -> Option<SrcRect> {
        let inner = clip.trim().strip_prefix("rect(")?.strip_suffix(')')?;
        let values: Vec<f64> = inner
            .split([',', ' '])
            .filter(|p| !p.is_empty())
            .map(|p| if p == "auto" { Ok(0.0) } else { length_pt(p) })
            .collect::<Result<_, _>>()
            .ok()?;
        if values.len() != 4 || values.iter().all(|v| *v == 0.0) {
            return None;
        }
        let (width, height) = self.intrinsic_size(path)?;
        if width <= 0.0 || height <= 0.0 {
            return None;
        }
        Some(SrcRect {
            t: values[0] / height,
            r: values[1] / width,
            b: values[2] / height,
            l: values[3] / width,
        })
    }

    /// A text frame (§10.4.3 `draw:text-box`) as a DOCX text box shape.
    #[allow(clippy::too_many_arguments)]
    fn text_box(
        &mut self,
        frame: roxmltree::Node<'a, 'a>,
        text_box: roxmltree::Node<'a, 'a>,
        graphic: &Graphic<'a>,
        placement: Placement,
        width: f64,
        height: f64,
        rotation: f64,
    ) -> Result<Option<DocRun>, String> {
        let padding = |side: &str| {
            graphic
                .length(NS_FO, &format!("padding-{side}"))
                .or_else(|| graphic.length(NS_FO, "padding"))
                .unwrap_or(0.0)
                .max(0.0)
        };
        let insets = [
            padding("left"),
            padding("top"),
            padding("right"),
            padding("bottom"),
        ];
        // One border for the box: the first side that has one (the
        // shorthand is written out as its sides).
        let border = ["top", "left", "bottom", "right"]
            .into_iter()
            .find_map(|side| {
                graphic
                    .get(NS_FO, &format!("border-{side}"))
                    .and_then(|value| parse_border(value, None))
            });
        let border_width = border.as_ref().map_or(0.0, |b| b.width);
        let inner_width = (width - insets[0] - insets[2] - 2.0 * border_width).max(1.0);
        // The frame's paragraphs continue their property search in its
        // graphic style, like a drawing shape's (§16.2).
        let content = self.contained_story_in(text_box, inner_width, graphic.chain.clone())?;
        let grows = frame.attribute((NS_FO, "min-height")).is_some()
            || text_box.attribute((NS_FO, "min-height")).is_some()
            || graphic.get(NS_DRAW, "auto-grow-height") == Some("true");
        let fill = self.fill(graphic, true)?;
        let mut shape = ShapeRun {
            inline: placement.inline,
            width_pt: width,
            height_pt: height,
            preset_geometry: Some("rect".into()),
            fill,
            fill_opacity: graphic
                .get(NS_DRAW, "opacity")
                .and_then(|v| percent(v).ok())
                .filter(|v| *v < 1.0),
            stroke: border
                .as_ref()
                .map(|b| b.color.clone().unwrap_or_else(|| "000000".into())),
            stroke_width: border_width,
            rotation,
            text_box_content: content.into_iter().map(TextBoxBlockWire::Body).collect(),
            text_anchor: Some(
                match graphic.get(NS_DRAW, "textarea-vertical-align") {
                    Some("middle") => "ctr",
                    Some("bottom") => "b",
                    _ => "t",
                }
                .into(),
            ),
            text_autofit: Some(if grows { "sp" } else { "none" }.into()),
            text_inset_l: insets[0] + border_width,
            text_inset_t: insets[1] + border_width,
            text_inset_r: insets[2] + border_width,
            text_inset_b: insets[3] + border_width,
            text_vert: match graphic.get(NS_STYLE, "writing-mode") {
                Some("tb-rl" | "tb") => Some("eaVert".into()),
                _ => None,
            },
            ..ShapeRun::default()
        };
        if !placement.inline {
            self.anchor_shape(&mut shape, &placement);
        }
        Ok(Some(DocRun::Shape(Box::new(shape))))
    }

    fn anchor_shape(&mut self, shape: &mut ShapeRun, placement: &Placement) {
        let acquisition = self.acquisition(placement, shape.width_pt, shape.height_pt);
        if let AnchorAxisChoiceWire::Offset { value_pt } = placement.horizontal.1 {
            shape.anchor_x_pt = value_pt;
        }
        if let AnchorAxisChoiceWire::Offset { value_pt } = placement.vertical.1 {
            shape.anchor_y_pt = value_pt;
        }
        if let AnchorAxisChoiceWire::Align { value } = &placement.horizontal.1 {
            shape.anchor_x_align = Some(value.clone());
        }
        if let AnchorAxisChoiceWire::Align { value } = &placement.vertical.1 {
            shape.anchor_y_align = Some(value.clone());
        }
        shape.anchor_x_from_margin = placement.horizontal.0 != "page";
        shape.anchor_y_from_para = placement.vertical.0 == "paragraph";
        shape.anchor_x_relative_from = Some(placement.horizontal.0.clone());
        shape.anchor_y_relative_from = Some(placement.vertical.0.clone());
        shape.behind_doc = placement.behind;
        shape.z_order = placement.z_order;
        shape.wrap_mode = Some(wrap_mode(placement.effective_wrap()).into());
        shape.wrap_side = placement.side.map(str::to_string);
        shape.dist_top = placement.distances[0];
        shape.dist_right = placement.distances[1];
        shape.dist_bottom = placement.distances[2];
        shape.dist_left = placement.distances[3];
        shape.anchor_acquisition = Some(acquisition);
    }

    /// A page background's fill from its captured properties.
    pub(super) fn background_fill(
        &mut self,
        style: &'a crate::style::Style,
    ) -> Result<Option<ShapeFill>, String> {
        let style = StyleRef::new(style);
        let graphic = Graphic {
            cascade: Cascade::build(vec![style], "graphic-properties")?,
            chain: Rc::from(Vec::new()),
        };
        self.fill(&graphic, true)
    }

    /// A behind-text shape covering the page, anchored to the page.
    pub(super) fn page_background_shape(
        &mut self,
        fill: ShapeFill,
        width: f64,
        height: f64,
    ) -> DocRun {
        let placement = Placement {
            inline: false,
            horizontal: (
                "page".into(),
                AnchorAxisChoiceWire::Offset { value_pt: 0.0 },
            ),
            vertical: (
                "page".into(),
                AnchorAxisChoiceWire::Offset { value_pt: 0.0 },
            ),
            wrap: AnchorWrapKindWire::None,
            side: None,
            behind: true,
            distances: [0.0; 4],
            z_order: 0,
            allow_overlap: true,
            polygon: None,
        };
        let mut shape = ShapeRun {
            width_pt: width,
            height_pt: height,
            preset_geometry: Some("rect".into()),
            fill: Some(fill),
            ..ShapeRun::default()
        };
        self.anchor_shape(&mut shape, &placement);
        DocRun::Shape(Box::new(shape))
    }

    /// The area fill of a graphic style (§20.120 draw:fill and its
    /// companions), or of a frame's `fo:background-color`. A gradient
    /// definition is shared, but each fill owns its stops.
    fn fill(&mut self, graphic: &Graphic<'a>, frame: bool) -> Result<Option<ShapeFill>, String> {
        let color = |value: Option<&str>| value.and_then(|v| hex_color(v).ok());
        let definition = |reader: &Self, kind: &'static str, attribute: &str| match graphic
            .get(NS_DRAW, attribute)
        {
            Some(name) => reader.definition(kind, name),
            None => Ok(None),
        };
        Ok(match graphic.get(NS_DRAW, "fill") {
            Some("none") => None,
            Some("solid") => {
                color(graphic.get(NS_DRAW, "fill-color")).map(|color| ShapeFill::Solid { color })
            }
            Some("gradient") => definition(self, "gradient", "fill-gradient-name")?.map(gradient),
            Some("hatch") => {
                let hatch = definition(self, "hatch", "fill-hatch-name")?;
                let fg = hatch
                    .and_then(|h| color(h.attribute((NS_DRAW, "color"))))
                    .unwrap_or_else(|| "000000".into());
                let bg = if graphic.get(NS_DRAW, "fill-hatch-solid") == Some("true") {
                    color(graphic.get(NS_DRAW, "fill-color")).unwrap_or_else(|| "ffffff".into())
                } else {
                    "ffffff".into()
                };
                let preset = match hatch.and_then(|h| h.attribute((NS_DRAW, "style"))) {
                    Some("double") => "smGrid",
                    Some("triple") => "diagCross",
                    _ => "ltUpDiag",
                };
                Some(ShapeFill::Pattern {
                    fg,
                    bg,
                    preset: preset.into(),
                })
            }
            Some("bitmap") => {
                let href = definition(self, "fill-image", "fill-image-name")?
                    .and_then(|image| image.attribute((NS_XLINK, "href")));
                let Some(href) = href else {
                    return Ok(None);
                };
                let Some((path, mime)) = self.package_image(href)? else {
                    return Ok(None);
                };
                let tile = (graphic.get(NS_STYLE, "repeat") == Some("repeat")).then(|| {
                    Box::new(ooxml_common::fill::TileInfo {
                        tx: None,
                        ty: None,
                        sx: None,
                        sy: None,
                        flip: None,
                        algn: None,
                    })
                });
                Some(ShapeFill::Image {
                    image_path: path,
                    mime_type: mime.into(),
                    svg_image_path: None,
                    src_rect: None,
                    fill_rect: None,
                    tile,
                    alpha: None,
                    duotone: None,
                    blip_effects: Vec::new(),
                    vml: None,
                })
            }
            _ if frame => match graphic.get(NS_FO, "background-color") {
                Some("transparent") | None => None,
                Some(value) => hex_color(value)
                    .ok()
                    .map(|color| ShapeFill::Solid { color }),
            },
            // Shapes default to LibreOffice's default graphic fill when the
            // style chain sets one.
            _ => color(graphic.get(NS_DRAW, "fill-color")).map(|color| ShapeFill::Solid { color }),
        })
    }

    /// A drawing shape (§10.3, §10.6) as a DOCX shape with its geometry,
    /// fill, line and text.
    fn shape(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        group_offset: Option<(f64, f64)>,
    ) -> Result<Option<DocRun>, String> {
        let graphic = self.graphic(node)?;
        let get = |name: &str| {
            node.attribute((NS_SVG, name))
                .and_then(|v| length_pt(v).ok())
        };
        let name = node.tag_name().name();
        // Lines and connectors: their two points.
        let line_points = matches!(name, "line" | "connector" | "measure").then(|| {
            (
                get("x1").unwrap_or(0.0),
                get("y1").unwrap_or(0.0),
                get("x2").unwrap_or(0.0),
                get("y2").unwrap_or(0.0),
            )
        });
        let (mut width, mut height) = (get("width").unwrap_or(0.0), get("height").unwrap_or(0.0));
        let mut origin = (get("x").unwrap_or(0.0), get("y").unwrap_or(0.0));
        if let Some((x1, y1, x2, y2)) = line_points {
            origin = (x1.min(x2), y1.min(y2));
            width = (x2 - x1).abs();
            height = (y2 - y1).abs();
        }
        let (rotation, translated) = transform_rotation(
            node.attribute((NS_DRAW, "transform")),
            origin,
            width,
            height,
        );
        // The geometry of a line is relative to its untransformed box.
        let box_origin = origin;
        if let Some(position) = translated {
            origin = position;
        }
        // Group members carry absolute positions within the group's frame.
        if let Some(offset) = group_offset {
            origin = (origin.0 + offset.0, origin.1 + offset.1);
        }
        let mut placement = self.placement(node, &graphic, (0.0, 0.0))?;
        placement.horizontal.1 = match placement.horizontal.1 {
            AnchorAxisChoiceWire::Offset { .. } => {
                AnchorAxisChoiceWire::Offset { value_pt: origin.0 }
            }
            other => other,
        };
        placement.vertical.1 = match placement.vertical.1 {
            AnchorAxisChoiceWire::Offset { .. } => {
                AnchorAxisChoiceWire::Offset { value_pt: origin.1 }
            }
            other => other,
        };
        let mut shape = ShapeRun {
            inline: placement.inline,
            width_pt: width.max(0.0),
            height_pt: height.max(0.0),
            rotation,
            ..ShapeRun::default()
        };
        // Geometry.
        match name {
            "rect" => {
                let radius = node
                    .attribute((NS_DRAW, "corner-radius"))
                    .and_then(|v| length_pt(v).ok())
                    .unwrap_or(0.0);
                if radius > 0.0 && width.min(height) > 0.0 {
                    shape.preset_geometry = Some("roundRect".into());
                    shape.adj_values =
                        vec![Some((radius / width.min(height) * 100_000.0).min(50_000.0))];
                } else {
                    shape.preset_geometry = Some("rect".into());
                }
            }
            "ellipse" | "circle" => shape.preset_geometry = Some("ellipse".into()),
            "line" | "connector" | "measure" => {
                let (x1, y1, x2, y2) = line_points.unwrap_or_default();
                let n = |v: f64, origin: f64, size: f64| {
                    if size > 0.0 {
                        (v - origin) / size
                    } else {
                        0.0
                    }
                };
                shape.subpaths = vec![vec![
                    PathCmd::MoveTo {
                        x: n(x1, box_origin.0, width),
                        y: n(y1, box_origin.1, height),
                    },
                    PathCmd::LineTo {
                        x: n(x2, box_origin.0, width),
                        y: n(y2, box_origin.1, height),
                    },
                ]];
            }
            "polyline" | "polygon" => {
                let view = view_box(node).unwrap_or((0.0, 0.0, width.max(1.0), height.max(1.0)));
                let source = node.attribute((NS_DRAW, "points")).unwrap_or("");
                let points: Vec<(f64, f64)> = source
                    .split_whitespace()
                    .filter_map(|pair| {
                        let (x, y) = pair.split_once(',')?;
                        Some((x.parse().ok()?, y.parse().ok()?))
                    })
                    .collect();
                budget::tick(1)?;
                let mut path = Vec::new();
                for (index, (x, y)) in points.iter().enumerate() {
                    let x = if view.2 != 0.0 {
                        (x - view.0) / view.2
                    } else {
                        0.0
                    };
                    let y = if view.3 != 0.0 {
                        (y - view.1) / view.3
                    } else {
                        0.0
                    };
                    path.push(if index == 0 {
                        PathCmd::MoveTo { x, y }
                    } else {
                        PathCmd::LineTo { x, y }
                    });
                }
                if name == "polygon" && !path.is_empty() {
                    path.push(PathCmd::Close);
                }
                if !path.is_empty() {
                    shape.subpaths = vec![path];
                }
            }
            "path" => {
                let view = view_box(node).unwrap_or((0.0, 0.0, width.max(1.0), height.max(1.0)));
                if let Some(d) = node.attribute((NS_SVG, "d")) {
                    shape.subpaths = svg_path(d, view)?;
                }
            }
            "custom-shape" => {
                if let Some(geometry) = node
                    .kids()
                    .find(|n| xml::is(*n, NS_DRAW, "enhanced-geometry"))
                {
                    self.enhanced_geometry(geometry, &graphic, &mut shape, width, height)?;
                } else {
                    shape.preset_geometry = Some("rect".into());
                }
            }
            _ => shape.preset_geometry = Some("rect".into()),
        }
        // Fill and line.
        let closed = !matches!(name, "line" | "connector" | "measure" | "polyline");
        shape.fill = if closed {
            self.fill(&graphic, false)?
        } else {
            None
        };
        shape.fill_opacity = graphic
            .get(NS_DRAW, "opacity")
            .and_then(|v| percent(v).ok())
            .filter(|v| *v < 1.0);
        self.line(&graphic, &mut shape)?;
        // Text in the shape (§10.6.1: paragraphs and lists).
        if node
            .kids()
            .any(|n| xml::is(n, NS_TEXT, "p") || xml::is(n, NS_TEXT, "list"))
        {
            let padding = |side: &str| {
                graphic
                    .length(NS_FO, &format!("padding-{side}"))
                    .or_else(|| graphic.length(NS_FO, "padding"))
                    .unwrap_or(0.0)
                    .max(0.0)
            };
            let insets = [
                padding("left"),
                padding("top"),
                padding("right"),
                padding("bottom"),
            ];
            let inner = (width - insets[0] - insets[2]).max(1.0);
            // The paragraphs of a drawing shape continue their property
            // search in the shape's graphic style (§16.2).
            let content = self.contained_story_in(node, inner, graphic.chain.clone())?;
            shape.text_box_content = content.into_iter().map(TextBoxBlockWire::Body).collect();
            shape.text_inset_l = insets[0];
            shape.text_inset_t = insets[1];
            shape.text_inset_r = insets[2];
            shape.text_inset_b = insets[3];
            shape.text_anchor = Some(
                match graphic.get(NS_DRAW, "textarea-vertical-align") {
                    Some("top") => "t",
                    Some("bottom") => "b",
                    _ => "ctr",
                }
                .into(),
            );
            shape.text_autofit = Some(
                if graphic.get(NS_DRAW, "auto-grow-height") == Some("true") {
                    "sp"
                } else {
                    "none"
                }
                .into(),
            );
        }
        if shape.subpaths.is_empty() && shape.preset_geometry.is_none() {
            return Ok(None);
        }
        if !placement.inline {
            self.anchor_shape(&mut shape, &placement);
        }
        Ok(Some(DocRun::Shape(Box::new(shape))))
    }

    /// `draw:enhanced-geometry`: the evaluated path, or a preset for the
    /// shape types without one.
    fn enhanced_geometry(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        graphic: &Graphic<'a>,
        shape: &mut ShapeRun,
        width: f64,
        height: f64,
    ) -> Result<(), String> {
        let kind = node.attribute((NS_DRAW, "type")).unwrap_or("non-primitive");
        shape.flip_h = node.attribute((NS_DRAW, "mirror-horizontal")) == Some("true");
        shape.flip_v = node.attribute((NS_DRAW, "mirror-vertical")) == Some("true");
        if let Some(path) = node.attribute((NS_DRAW, "enhanced-path")) {
            // `draw:modifiers` (§19.196): the numbers `$n` refers to.
            let mut modifiers: Vec<f64> = Vec::new();
            for value in node
                .attribute((NS_DRAW, "modifiers"))
                .unwrap_or("")
                .split_whitespace()
            {
                budget::tick(1)?;
                modifiers.extend(value.parse::<f64>().ok());
            }
            // `draw:equation` (§10.6.3) definitions, within the reader's
            // bound: a shape with more is rejected, not cut short.
            let mut equations = HashMap::new();
            for equation in node.kids().filter(|n| xml::is(*n, NS_DRAW, "equation")) {
                budget::tick(1)?;
                if equations.len() >= geometry::MAX_EQUATIONS {
                    return Err(resource("shape equation"));
                }
                if let (Some(name), Some(formula)) = (
                    equation.attribute((NS_DRAW, "name")),
                    equation.attribute((NS_DRAW, "formula")),
                ) {
                    equations.insert(name.to_string(), formula.to_string());
                }
            }
            let geometry = geometry::Geometry {
                view_box: view_box(node).unwrap_or((0.0, 0.0, 21600.0, 21600.0)),
                equations,
                modifiers,
                log_size: (width * 2540.0 / 72.0, height * 2540.0 / 72.0),
                has_fill: graphic.get(NS_DRAW, "fill") != Some("none"),
                has_stroke: graphic.get(NS_DRAW, "stroke") != Some("none"),
            };
            let path = geometry::evaluate(path, &geometry)?;
            if !path.subpaths.is_empty() {
                shape.subpaths = path.subpaths;
                if path.paint.iter().any(|p| p.fill.is_some() || !p.stroke) {
                    shape.subpath_paint = path.paint;
                }
                return Ok(());
            }
        }
        shape.preset_geometry = Some(
            match kind.strip_prefix("ooxml-") {
                Some(preset) => preset,
                None => match kind {
                    "ellipse" => "ellipse",
                    "round-rectangle" => "roundRect",
                    "isosceles-triangle" => "triangle",
                    "right-triangle" => "rtTriangle",
                    "diamond" => "diamond",
                    "hexagon" => "hexagon",
                    "octagon" => "octagon",
                    "parallelogram" => "parallelogram",
                    "trapezoid" => "trapezoid",
                    "pentagon" => "pentagon",
                    _ => "rect",
                },
            }
            .to_string(),
        );
        Ok(())
    }

    /// The line of a shape (§20.143 draw:stroke and companions).
    fn line(&mut self, graphic: &Graphic<'a>, shape: &mut ShapeRun) -> Result<(), String> {
        let stroke = graphic.get(NS_DRAW, "stroke").unwrap_or("solid");
        if stroke == "none" {
            return Ok(());
        }
        shape.stroke = Some(
            graphic
                .get(NS_SVG, "stroke-color")
                .and_then(|v| hex_color(v).ok())
                .unwrap_or_else(|| "000000".into()),
        );
        // A zero width is a hairline.
        shape.stroke_width = graphic
            .length(NS_SVG, "stroke-width")
            .filter(|w| *w > 0.0)
            .unwrap_or(0.5);
        if stroke == "dash" {
            let dash = match graphic.get(NS_DRAW, "stroke-dash") {
                Some(name) => self.definition("stroke-dash", name)?,
                None => None,
            };
            let dotted = dash.is_some_and(|d| {
                d.attribute((NS_DRAW, "dots1-length"))
                    .and_then(|v| v.strip_suffix('%').and_then(|n| n.parse::<f64>().ok()))
                    .is_some_and(|v| v <= 100.0)
                    && d.attribute((NS_DRAW, "dots2")).is_none()
            });
            shape.stroke_dash = Some(if dotted { "sysDot" } else { "dash" }.into());
        }
        shape.stroke_join = graphic.get(NS_DRAW, "stroke-linejoin").map(|join| {
            match join {
                "round" => "round",
                "bevel" => "bevel",
                _ => "miter",
            }
            .into()
        });
        shape.stroke_cap = graphic.get(NS_SVG, "stroke-linecap").map(|cap| {
            match cap {
                "round" => "rnd",
                "square" => "sq",
                _ => "flat",
            }
            .into()
        });
        let marker = |name: Option<&str>, width: Option<f64>| -> Option<LineEnd> {
            let name = name?;
            let kind = if name.contains("Circle") {
                "oval"
            } else if name.contains("Square") || name.contains("Diamond") {
                "diamond"
            } else if name.contains("Line") || name.contains("heads") {
                "arrow"
            } else {
                "triangle"
            };
            let ratio = width.unwrap_or(0.0) / shape.stroke_width.max(0.5);
            let size = if ratio <= 2.5 {
                "sm"
            } else if ratio <= 4.5 {
                "med"
            } else {
                "lg"
            };
            Some(LineEnd {
                r#type: kind.into(),
                w: size.into(),
                len: size.into(),
            })
        };
        shape.head_end = marker(
            graphic.get(NS_DRAW, "marker-start"),
            graphic.length(NS_DRAW, "marker-start-width"),
        );
        shape.tail_end = marker(
            graphic.get(NS_DRAW, "marker-end"),
            graphic.length(NS_DRAW, "marker-end-width"),
        );
        Ok(())
    }

    /// `draw:g` (§10.3.15): its members as separate shapes on the group's
    /// anchor (DOCX anchors hold one object each).
    fn group(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<Option<DocRun>, String> {
        let graphic = self.graphic(node)?;
        // §19.759 values only (a validated, fixed-size anchor). An
        // as-character group cannot stay one inline object in the DOCX
        // model; its members keep their paragraph anchoring.
        // The outermost group is the anchored object: an enclosing group's
        // anchor wins over a nested group's own.
        let anchor = self.group_anchor.or_else(|| {
            match node
                .attribute((NS_TEXT, "anchor-type"))
                .or_else(|| graphic.get(NS_TEXT, "anchor-type"))
            {
                Some("page") => Some("page"),
                Some("paragraph") => Some("paragraph"),
                Some("char") => Some("char"),
                Some("frame") => Some("frame"),
                // An as-character group keeps its members' own anchoring.
                Some("as-char") => None,
                // §19.759: an omitted (or unknown) anchor type is paragraph.
                _ => Some("paragraph"),
            }
        });
        let saved = std::mem::replace(&mut self.group_anchor, anchor);
        self.group_depth += 1;
        let result = self.group_members(node);
        self.group_depth -= 1;
        self.group_anchor = saved;
        result
    }

    fn group_members(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<Option<DocRun>, String> {
        let mut runs = Vec::new();
        for child in node.kids().filter(|n| n.is_element()) {
            budget::tick(1)?;
            // Through `drawing`, which bounds the nesting depth.
            let run = self.drawing(child, None)?;
            if let Some(run) = run {
                runs.push(run);
            }
        }
        // The first member here; the rest join the paragraph's pending
        // anchors.
        let mut runs = runs.into_iter();
        let first = runs.next();
        self.pending_anchors.extend(runs);
        Ok(first)
    }
}

/// `draw:contour-polygon` points in its view box, scaled to DOCX's
/// 21600-unit wrap polygon space.
fn contour_polygon(
    frame: roxmltree::Node<'_, '_>,
) -> Result<Option<docx_model::AnchorPolygonWire>, String> {
    let Some(contour) = frame
        .kids()
        .find(|n| xml::is(*n, NS_DRAW, "contour-polygon"))
    else {
        return Ok(None);
    };
    let (Some(view), Some(source)) = (view_box(contour), contour.attribute((NS_DRAW, "points")))
    else {
        return Ok(None);
    };
    budget::tick(1)?;
    let points: Vec<docx_model::AnchorPointWire> = source
        .split_whitespace()
        .filter_map(|pair| {
            let (x, y) = pair.split_once(',')?;
            let (x, y): (f64, f64) = (x.parse().ok()?, y.parse().ok()?);
            Some(docx_model::AnchorPointWire {
                x: Some(((x - view.0) / view.2 * 21600.0).round() as i64),
                y: Some(((y - view.1) / view.3 * 21600.0).round() as i64),
                raw_x: None,
                raw_y: None,
            })
        })
        .collect();
    Ok(
        (points.len() >= 3).then_some(docx_model::AnchorPolygonWire {
            edited: false,
            coordinate_space: docx_model::AnchorPolygonSpaceWire {
                width: 21600,
                height: 21600,
            },
            points,
            invalid_point_count: 0,
        }),
    )
}

fn wrap_mode(kind: AnchorWrapKindWire) -> &'static str {
    match kind {
        AnchorWrapKindWire::TopAndBottom => "topAndBottom",
        AnchorWrapKindWire::Square => "square",
        AnchorWrapKindWire::Tight => "tight",
        AnchorWrapKindWire::Through => "through",
        _ => "none",
    }
}

/// `svg:viewBox` (§19.574).
fn view_box(node: roxmltree::Node<'_, '_>) -> Option<(f64, f64, f64, f64)> {
    let values: Vec<f64> = node
        .attribute((NS_SVG, "viewBox"))?
        .split([' ', ','])
        .filter(|v| !v.is_empty())
        .filter_map(|v| v.parse().ok())
        .collect();
    (values.len() == 4 && values[2] != 0.0 && values[3] != 0.0)
        .then(|| (values[0], values[1], values[2], values[3]))
}

/// `draw:transform` (§19.228): the list of transforms composed in order
/// into one affine map of the unrotated box at `origin`, read as the box's
/// rotation (DOCX clockwise degrees) and, where the list holds a transform,
/// the box's top-left corner (the map of its centre, less half its size).
/// Scale and skew move the centre; the box keeps its size.
///
/// Angles are radians, counter-clockwise on the page: what LibreOffice
/// writes and reads (every rotated object of the corpus, for example
/// `rotate (-1.5707963267949)` for a quarter turn). The specification's
/// text says degrees; reading them so would turn every LibreOffice file.
fn transform_rotation(
    value: Option<&str>,
    origin: (f64, f64),
    width: f64,
    height: f64,
) -> (f64, Option<(f64, f64)>) {
    let Some(value) = value else {
        return (0.0, None);
    };
    // x' = a x + c y + e, y' = b x + d y + f.
    let mut map = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
    let mut positioned = false;
    let mut rest = value;
    while let Some(open) = rest.find('(') {
        let name = rest[..open].trim_matches(|c: char| c.is_whitespace() || c == ',');
        let Some(close) = rest[open..].find(')') else {
            break;
        };
        let mut arguments = rest[open + 1..open + close]
            .split([' ', ','])
            .filter(|v| !v.is_empty())
            .filter_map(|v| length_pt(v).ok().or_else(|| v.parse::<f64>().ok()));
        let mut next = |default: f64| arguments.next().unwrap_or(default);
        let step = match name {
            "rotate" => {
                let (s, c) = next(0.0).sin_cos();
                Some([c, -s, s, c, 0.0, 0.0])
            }
            "translate" => Some([1.0, 0.0, 0.0, 1.0, next(0.0), next(0.0)]),
            "scale" => {
                let sx = next(1.0);
                Some([sx, 0.0, 0.0, next(sx), 0.0, 0.0])
            }
            // The shear LibreOffice draws (shared with the other readers).
            "skewX" => Some([
                1.0,
                0.0,
                crate::odf_props::skew_shear(next(0.0)),
                1.0,
                0.0,
                0.0,
            ]),
            "skewY" => Some([
                1.0,
                crate::odf_props::skew_shear(next(0.0)),
                0.0,
                1.0,
                0.0,
                0.0,
            ]),
            "matrix" => Some([
                next(1.0),
                next(0.0),
                next(0.0),
                next(1.0),
                next(0.0),
                next(0.0),
            ]),
            _ => None,
        };
        if let Some([a, b, c, d, e, f]) = step {
            positioned = true;
            // This transform after the ones before it.
            let [pa, pb, pc, pd, pe, pf] = map;
            map = [
                a * pa + c * pb,
                b * pa + d * pb,
                a * pc + c * pd,
                b * pc + d * pd,
                a * pe + c * pf + e,
                b * pe + d * pf + f,
            ];
        }
        rest = &rest[open + close + 1..];
    }
    if map.iter().any(|v| !v.is_finite()) {
        return (0.0, None);
    }
    let [a, b, c, d, e, f] = map;
    // The angle of the box's x axis (`rotate(t)` maps it to (cos t, -sin t)).
    let angle = (-b).atan2(a);
    let degrees = (-angle.to_degrees()).rem_euclid(360.0);
    let corner = positioned.then(|| {
        let (hx, hy) = (width / 2.0, height / 2.0);
        let (cx, cy) = (origin.0 + hx, origin.1 + hy);
        (a * cx + c * cy + e - hx, b * cx + d * cy + f - hy)
    });
    (degrees, corner)
}

/// A gradient definition (§16.40) as a DOCX gradient fill.
fn gradient(node: roxmltree::Node<'_, '_>) -> ShapeFill {
    let color = |name: &str, default: &str| {
        node.attribute((NS_DRAW, name))
            .and_then(|v| hex_color(v).ok())
            .unwrap_or_else(|| default.to_string())
    };
    let start = color("start-color", "000000");
    let end = color("end-color", "ffffff");
    // §19.218 draw:angle: tenths of a degree without a unit, counter-
    // clockwise; 0 runs top to bottom.
    let angle = node
        .attribute((NS_DRAW, "angle"))
        .map(|v| {
            if let Some(deg) = v.strip_suffix("deg") {
                deg.parse::<f64>().unwrap_or(0.0)
            } else if let Some(rad) = v.strip_suffix("rad") {
                rad.parse::<f64>().unwrap_or(0.0).to_degrees()
            } else {
                v.parse::<f64>().unwrap_or(0.0) / 10.0
            }
        })
        .unwrap_or(0.0);
    let style = node.attribute((NS_DRAW, "style")).unwrap_or("linear");
    let mut stops: Vec<GradientStop> = node
        .kids()
        .filter(|n| xml::is(*n, NS_LOEXT, "gradient-stop"))
        .filter_map(|stop| {
            Some(GradientStop {
                position: stop
                    .attribute((NS_SVG, "offset"))
                    .and_then(|v| v.parse::<f64>().ok())?,
                color: stop
                    .attribute((NS_LOEXT, "color-value"))
                    .and_then(|v| hex_color(v).ok())?,
            })
        })
        .collect();
    if stops.len() < 2 {
        stops = if style == "axial" {
            vec![
                GradientStop {
                    position: 0.0,
                    color: end.clone(),
                },
                GradientStop {
                    position: 0.5,
                    color: start,
                },
                GradientStop {
                    position: 1.0,
                    color: end,
                },
            ]
        } else {
            vec![
                GradientStop {
                    position: 0.0,
                    color: start,
                },
                GradientStop {
                    position: 1.0,
                    color: end,
                },
            ]
        };
    }
    ShapeFill::Gradient {
        stops,
        // DrawingML angles run clockwise from left-to-right.
        angle: (90.0 - angle).rem_euclid(360.0),
        grad_type: match style {
            "radial" | "ellipsoid" => "circle".into(),
            "square" | "rectangular" => "rect".into(),
            _ => "linear".into(),
        },
        scaled: None,
        path: match style {
            "radial" | "ellipsoid" => Some("circle".into()),
            "square" | "rectangular" => Some("rect".into()),
            _ => None,
        },
        fill_to_rect: None,
        tile_rect: None,
        flip: None,
        rot_with_shape: None,
    }
}

/// An SVG path (`svg:d`, §19.554: SVG 1.1 §8.3 path data) in the view box,
/// normalized: absolute and relative M, L, H, V, C, S, Q, T, A and Z. `S` and
/// `T` take their first control point as the reflection of the previous
/// curve's last one about the current point (the current point itself after
/// any other command; SVG 1.1 §8.3.6, §8.3.7). `A` is the elliptical arc of
/// SVG 1.1 §F.6 (§F.6.5 endpoint-to-centre conversion, §F.6.6 out-of-range
/// radii scaled up, §F.6.2 zero radii a line), drawn as cubic Béziers of at
/// most a quarter turn.
fn svg_path(d: &str, view: (f64, f64, f64, f64)) -> Result<Vec<Vec<PathCmd>>, String> {
    let norm = |(x, y): (f64, f64)| ((x - view.0) / view.2, (y - view.1) / view.3);
    let mut reader = SvgData {
        text: d.as_bytes(),
        at: 0,
    };
    let mut paths: Vec<Vec<PathCmd>> = Vec::new();
    let mut path: Vec<PathCmd> = Vec::new();
    let mut point = (0.0, 0.0);
    let mut start = (0.0, 0.0);
    let mut command: Option<u8> = None;
    // The previous command's last control point, for S (after C / S) and T
    // (after Q / T).
    let mut last_cubic: Option<(f64, f64)> = None;
    let mut last_quad: Option<(f64, f64)> = None;
    let line = |path: &mut Vec<PathCmd>, p: (f64, f64)| {
        let (x, y) = norm(p);
        path.push(PathCmd::LineTo { x, y });
    };
    let cubic = |path: &mut Vec<PathCmd>, c1: (f64, f64), c2: (f64, f64), p: (f64, f64)| {
        let ((x1, y1), (x2, y2), (x, y)) = (norm(c1), norm(c2), norm(p));
        path.push(PathCmd::CubicBezTo {
            x1,
            y1,
            x2,
            y2,
            x,
            y,
        });
    };
    let reflect = |control: Option<(f64, f64)>, about: (f64, f64)| {
        control.map_or(about, |(x, y)| (2.0 * about.0 - x, 2.0 * about.1 - y))
    };
    loop {
        // Each step adds at most one command (an arc at most four).
        budget::tick(1)?;
        reader.skip();
        let Some(next) = reader.peek() else {
            break;
        };
        if next.is_ascii_alphabetic() {
            reader.at += 1;
            command = Some(next);
        }
        let Some(letter) = command else {
            break;
        };
        let base = if letter.is_ascii_lowercase() {
            point
        } else {
            (0.0, 0.0)
        };
        let abs = |x: f64, y: f64| (base.0 + x, base.1 + y);
        let (previous_cubic, previous_quad) = (last_cubic.take(), last_quad.take());
        let ok = match letter.to_ascii_uppercase() {
            b'Z' => {
                path.push(PathCmd::Close);
                point = start;
                // Numbers after Z without a command are an error (§8.3.3).
                command = None;
                true
            }
            b'M' => reader.numbers::<2>().is_some_and(|[x, y]| {
                if path.len() > 1 {
                    paths.push(std::mem::take(&mut path));
                } else {
                    path.clear();
                }
                point = abs(x, y);
                start = point;
                let (x, y) = norm(point);
                path.push(PathCmd::MoveTo { x, y });
                // Further pairs are implicit line-tos.
                command = Some(if letter == b'm' { b'l' } else { b'L' });
                true
            }),
            b'L' => reader.numbers::<2>().is_some_and(|[x, y]| {
                point = abs(x, y);
                line(&mut path, point);
                true
            }),
            b'H' => reader.numbers::<1>().is_some_and(|[x]| {
                point.0 = base.0 + x;
                line(&mut path, point);
                true
            }),
            b'V' => reader.numbers::<1>().is_some_and(|[y]| {
                point.1 = base.1 + y;
                line(&mut path, point);
                true
            }),
            b'C' => reader.numbers::<6>().is_some_and(|v| {
                let (c1, c2, p) = (abs(v[0], v[1]), abs(v[2], v[3]), abs(v[4], v[5]));
                cubic(&mut path, c1, c2, p);
                point = p;
                last_cubic = Some(c2);
                true
            }),
            b'S' => reader.numbers::<4>().is_some_and(|v| {
                let c1 = reflect(previous_cubic, point);
                let (c2, p) = (abs(v[0], v[1]), abs(v[2], v[3]));
                cubic(&mut path, c1, c2, p);
                point = p;
                last_cubic = Some(c2);
                true
            }),
            b'Q' | b'T' => {
                let quad = if letter.eq_ignore_ascii_case(&b'Q') {
                    reader
                        .numbers::<4>()
                        .map(|v| (abs(v[0], v[1]), abs(v[2], v[3])))
                } else {
                    reader
                        .numbers::<2>()
                        .map(|[x, y]| (reflect(previous_quad, point), abs(x, y)))
                };
                quad.is_some_and(|(c, p)| {
                    let ((x1, y1), (x, y)) = (norm(c), norm(p));
                    path.push(PathCmd::QuadBezTo { x1, y1, x, y });
                    point = p;
                    last_quad = Some(c);
                    true
                })
            }
            b'A' => reader
                .arc()
                .is_some_and(|(rx, ry, rotation, large, sweep, x, y)| {
                    let end = abs(x, y);
                    let curves = svg_arc(point, (rx, ry), rotation, large, sweep, end);
                    if curves.is_empty() && point != end {
                        // A zero radius (§F.6.2) or an arc beyond the arithmetic.
                        line(&mut path, end);
                    }
                    for (c1, c2, p) in curves {
                        cubic(&mut path, c1, c2, p);
                    }
                    point = end;
                    true
                }),
            _ => false,
        };
        if !ok {
            break;
        }
    }
    if path.len() > 1 {
        paths.push(path);
    }
    Ok(paths)
}

/// The SVG path-data reader: numbers (§8.3.9 grammar, the unsigned part read
/// by the formula grammar every reader shares) and arc flags.
struct SvgData<'t> {
    text: &'t [u8],
    at: usize,
}

impl SvgData<'_> {
    fn skip(&mut self) {
        while self
            .text
            .get(self.at)
            .is_some_and(|c| c.is_ascii_whitespace() || *c == b',')
        {
            self.at += 1;
        }
    }

    fn peek(&self) -> Option<u8> {
        self.text.get(self.at).copied()
    }

    fn number(&mut self) -> Option<f64> {
        self.skip();
        let start = self.at;
        let negative = match self.peek() {
            Some(b'-') => true,
            Some(b'+') => false,
            _ => {
                return crate::odf_props::scan_number(self.text, &mut self.at);
            }
        };
        self.at += 1;
        match crate::odf_props::scan_number(self.text, &mut self.at) {
            Some(value) => Some(if negative { -value } else { value }),
            None => {
                self.at = start;
                None
            }
        }
    }

    fn numbers<const N: usize>(&mut self) -> Option<[f64; N]> {
        let mut out = [0.0; N];
        for slot in &mut out {
            *slot = self.number()?;
        }
        Some(out)
    }

    /// An arc flag (`0` / `1`), which may abut the next number (§8.3.9).
    fn flag(&mut self) -> Option<bool> {
        self.skip();
        let value = match self.peek()? {
            b'0' => false,
            b'1' => true,
            _ => return None,
        };
        self.at += 1;
        Some(value)
    }

    #[allow(clippy::type_complexity)]
    fn arc(&mut self) -> Option<(f64, f64, f64, bool, bool, f64, f64)> {
        let [rx, ry, rotation] = self.numbers::<3>()?;
        let large = self.flag()?;
        let sweep = self.flag()?;
        let [x, y] = self.numbers::<2>()?;
        Some((rx, ry, rotation, large, sweep, x, y))
    }
}

type Cubic = ((f64, f64), (f64, f64), (f64, f64));

/// The cubic Béziers (control, control, end) of an SVG elliptical arc from
/// `from` to `to` (SVG 1.1 §F.6.5 endpoint-to-centre conversion, §F.6.6
/// radius correction), at most a quarter turn each. Empty for an arc that is
/// nothing or a line (§F.6.2), or beyond the range of the arithmetic.
fn svg_arc(
    from: (f64, f64),
    radii: (f64, f64),
    rotation: f64,
    large: bool,
    sweep: bool,
    to: (f64, f64),
) -> Vec<Cubic> {
    let (mut rx, mut ry) = (radii.0.abs(), radii.1.abs());
    if from == to || rx == 0.0 || ry == 0.0 {
        return Vec::new();
    }
    let (sin_phi, cos_phi) = rotation.to_radians().sin_cos();
    // Halves first, so neither the difference nor the midpoint overflows.
    let dx = from.0 / 2.0 - to.0 / 2.0;
    let dy = from.1 / 2.0 - to.1 / 2.0;
    let (mx, my) = (from.0 / 2.0 + to.0 / 2.0, from.1 / 2.0 + to.1 / 2.0);
    let x1p = cos_phi * dx + sin_phi * dy;
    let y1p = -sin_phi * dx + cos_phi * dy;
    // §F.6.5 in the unit-circle frame of the ellipse (`u`, `v` = x1', y1'
    // over the radii), with lengths taken by `hypot`: no square or product
    // of coordinates over- or underflows.
    let (mut u, mut v) = (x1p / rx, y1p / ry);
    let mut reach = u.hypot(v);
    if reach > 1.0 {
        // §F.6.6: radii too small for the ends grow until they fit.
        rx *= reach;
        ry *= reach;
        u /= reach;
        v /= reach;
        reach = 1.0;
    }
    // The centre in the frame is ±√(1−r²)·(v, −u)/r, r = |(u, v)|.
    let mut lift = ((1.0 - reach) * (1.0 + reach)).max(0.0).sqrt();
    if large == sweep {
        lift = -lift;
    }
    let (cu, cv) = (lift * (v / reach), -lift * (u / reach));
    let (cxp, cyp) = (cu * rx, cv * ry);
    let cx = cos_phi * cxp - sin_phi * cyp + mx;
    let cy = sin_phi * cxp + cos_phi * cyp + my;
    let angle = |ux: f64, uy: f64, vx: f64, vy: f64| (ux * vy - uy * vx).atan2(ux * vx + uy * vy);
    let theta1 = angle(1.0, 0.0, u - cu, v - cv);
    let mut delta = angle(u - cu, v - cv, -u - cu, -v - cv);
    if !sweep && delta > 0.0 {
        delta -= std::f64::consts::TAU;
    } else if sweep && delta < 0.0 {
        delta += std::f64::consts::TAU;
    }
    if ![cx, cy, rx, ry, theta1, delta]
        .iter()
        .all(|v| v.is_finite())
    {
        // Beyond the range of the arithmetic: the caller draws the chord.
        return Vec::new();
    }
    // The ellipse point and tangent at parameter t, rotated and placed.
    let place = |(x, y): (f64, f64)| (cos_phi * x - sin_phi * y, sin_phi * x + cos_phi * y);
    let at = |t: f64| {
        let (x, y) = place((rx * t.cos(), ry * t.sin()));
        (cx + x, cy + y)
    };
    let tangent = |t: f64| place((-rx * t.sin(), ry * t.cos()));
    // Quarter turns, not split by rounding (a quarter arc is one curve).
    let pieces = ((delta.abs() / std::f64::consts::FRAC_PI_2 - 1e-9).ceil() as usize).clamp(1, 4);
    let step = delta / pieces as f64;
    let k = 4.0 / 3.0 * (step / 4.0).tan();
    let curves: Vec<Cubic> = (0..pieces)
        .map(|i| {
            let (a, b) = (theta1 + step * i as f64, theta1 + step * (i + 1) as f64);
            let (p0, p1) = (at(a), at(b));
            let (ta, tb) = (tangent(a), tangent(b));
            let end = if i + 1 == pieces { to } else { p1 };
            (
                (p0.0 + k * ta.0, p0.1 + k * ta.1),
                (p1.0 - k * tb.0, p1.1 - k * tb.1),
                end,
            )
        })
        .collect();
    let finite = |(x, y): (f64, f64)| x.is_finite() && y.is_finite();
    if curves
        .iter()
        .all(|(c1, c2, p)| finite(*c1) && finite(*c2) && finite(*p))
    {
        curves
    } else {
        // A point beyond the range of the arithmetic: the caller draws the chord.
        Vec::new()
    }
}

#[cfg(test)]
mod path_tests {
    use super::*;

    const VIEW: (f64, f64, f64, f64) = (0.0, 0.0, 100.0, 100.0);

    /// The end point of every command of the only sub-path, in view units.
    fn ends(d: &str) -> Vec<(f64, f64)> {
        let paths = svg_path(d, VIEW).unwrap();
        assert_eq!(paths.len(), 1);
        paths[0]
            .iter()
            .filter_map(|c| match *c {
                PathCmd::MoveTo { x, y }
                | PathCmd::LineTo { x, y }
                | PathCmd::QuadBezTo { x, y, .. }
                | PathCmd::CubicBezTo { x, y, .. } => Some((x * 100.0, y * 100.0)),
                _ => None,
            })
            .collect()
    }

    fn close(actual: &[(f64, f64)], expected: &[(f64, f64)]) {
        assert_eq!(actual.len(), expected.len(), "{actual:?}");
        for (a, e) in actual.iter().zip(expected) {
            assert!(
                (a.0 - e.0).abs() < 1e-9 && (a.1 - e.1).abs() < 1e-9,
                "{actual:?}"
            );
        }
    }

    #[test]
    fn arcs_follow_the_svg_endpoint_parameterization() {
        // A half circle from (0,0) to (100,0): the sweep flag picks the side
        // (1 turns through positive angles, up the page here), each quarter
        // is one curve, and both spellings of the flags read alike.
        for d in [
            "M 0 0 A 50 50 0 0 1 100 0",
            "M0,0a50,50 0 0,1 100,0",
            "M0 0a50 50 0 01100 0",
        ] {
            close(&ends(d), &[(0.0, 0.0), (50.0, -50.0), (100.0, 0.0)]);
        }
        close(
            &ends("M 0 0 A 50 50 0 0 0 100 0"),
            &[(0.0, 0.0), (50.0, 50.0), (100.0, 0.0)],
        );
        // Radii too small for the ends are scaled up (§F.6.6).
        close(
            &ends("M 0 0 A 10 10 0 1 0 100 0"),
            &[(0.0, 0.0), (50.0, 50.0), (100.0, 0.0)],
        );
        // The x-axis rotation turns the ellipse: rx 20 upright, ry 50 flat.
        close(
            &ends("M 0 0 A 20 50 90 0 1 100 0"),
            &[(0.0, 0.0), (50.0, -20.0), (100.0, 0.0)],
        );
        // The large-arc flag takes the long way round: three quarters of
        // the circle about (50,-50) through the same ends.
        close(
            &ends("M 0 0 A 70.71067811865476 70.71067811865476 0 1 1 100 0"),
            &[(0.0, 0.0), (0.0, -100.0), (100.0, -100.0), (100.0, 0.0)],
        );
        // A zero radius is a line; equal ends draw nothing.
        close(&ends("M 0 0 A 0 5 0 0 1 10 0"), &[(0.0, 0.0), (10.0, 0.0)]);
        close(
            &ends("M 0 0 L 5 5 A 5 5 0 0 1 5 5"),
            &[(0.0, 0.0), (5.0, 5.0)],
        );
    }

    #[test]
    fn a_quarter_arc_is_the_standard_cubic() {
        // From (100,50) clockwise on the page to (50,100) about (50,50):
        // the controls sit k·r along the tangents, k = 4/3·tan(π/8).
        let paths = svg_path("M 100 50 A 50 50 0 0 1 50 100", VIEW).unwrap();
        let k = 4.0 / 3.0 * (std::f64::consts::PI / 8.0).tan() * 50.0;
        match paths[0][1] {
            PathCmd::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            } => {
                let got = [x1, y1, x2, y2, x, y].map(|v| v * 100.0);
                let want = [100.0, 50.0 + k, 50.0 + k, 100.0, 50.0, 100.0];
                for (g, w) in got.iter().zip(want) {
                    assert!((g - w).abs() < 1e-9, "{got:?}");
                }
            }
            ref other => panic!("{other:?}"),
        }
    }

    #[test]
    fn arcs_are_scale_invariant() {
        // A quarter circle about the origin, at tiny, unit and huge scales:
        // the same normalized controls (k = 4/3·tan(π/8) along the tangents).
        let k = 4.0 / 3.0 * (std::f64::consts::PI / 8.0).tan();
        for scale in [1e-100, 1.0, 1e100] {
            let d = format!("M{scale} 0 A{scale} {scale} 0 0 1 0 {scale}");
            let paths = svg_path(&d, (0.0, 0.0, scale, scale)).unwrap();
            let PathCmd::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            } = paths[0][1]
            else {
                panic!("{scale}: {:?}", paths[0]);
            };
            let want = [1.0, k, k, 1.0, 0.0, 1.0];
            for (got, want) in [x1, y1, x2, y2, x, y].iter().zip(want) {
                assert!((got - want).abs() < 1e-9, "{scale}: {:?}", paths[0][1]);
            }
        }
    }

    #[test]
    fn arcs_survive_extreme_finite_values() {
        // Tiny radii against a unit chord grow to the half circle (§F.6.6).
        close(
            &ends("M 0 0 A 1e-200 1e-200 0 0 1 100 0"),
            &[(0.0, 0.0), (50.0, -50.0), (100.0, 0.0)],
        );
        // Ends near the top of the range: their difference overflows unless
        // halved first. The half circle of radius 1.6e308 about the origin.
        let r = 1.6e308;
        let curves = svg_arc((-r, 0.0), (r, r), 0.0, false, true, (r, 0.0));
        assert_eq!(curves.len(), 2);
        let (_, _, top) = curves[0];
        assert!(
            (top.0 / r).abs() < 1e-12 && (top.1 / r + 1.0).abs() < 1e-12,
            "{top:?}"
        );
        assert_eq!(curves[1].2, (r, 0.0));
        // An arc whose points leave the finite range is drawn as its chord.
        let far = svg_arc(
            (1e308, 1e308),
            (1e308, 1e308),
            0.0,
            false,
            false,
            (1e308, -1e308),
        );
        assert!(far.is_empty());
        let paths = svg_path(
            "M 1e308 1e308 A 1e308 1e308 0 0 0 1e308 -1e308",
            (0.0, 0.0, 1.0, 1.0),
        )
        .unwrap();
        assert!(
            matches!(paths[0][1], PathCmd::LineTo { .. }),
            "{:?}",
            paths[0]
        );
    }

    #[test]
    fn smooth_curves_reflect_the_previous_control_point() {
        let paths = svg_path("M 0 0 C 0 10 20 10 20 0 S 40 -10 40 0", VIEW).unwrap();
        assert!(matches!(
            paths[0][2],
            PathCmd::CubicBezTo { x1, y1, x2, y2, x, y }
                if (x1, y1, x2, y2, x, y) == (0.2, -0.1, 0.4, -0.1, 0.4, 0.0)
        ));
        // Relative T after relative Q: the control mirrors (10,10) about (20,0).
        for d in ["M 0 0 Q 10 10 20 0 T 40 0", "m 0 0 q 10 10 20 0 t 20 0"] {
            let paths = svg_path(d, VIEW).unwrap();
            assert!(
                matches!(paths[0][2], PathCmd::QuadBezTo { x1, y1, x, y } if (x1, y1, x, y) == (0.3, -0.1, 0.4, 0.0)),
                "{d}: {:?}",
                paths[0][2]
            );
        }
        // After a non-curve command, S and T start at the current point.
        let paths = svg_path("M 0 0 L 10 0 S 20 10 30 0 L 40 0 T 50 0", VIEW).unwrap();
        assert!(matches!(
            paths[0][2],
            PathCmd::CubicBezTo { x1, y1, .. } if (x1, y1) == (0.1, 0.0)
        ));
        assert!(matches!(
            paths[0][4],
            PathCmd::QuadBezTo { x1, y1, .. } if (x1, y1) == (0.4, 0.0)
        ));
    }

    #[test]
    fn skew_moves_the_centre_as_libreoffice_shears() {
        // A 4 × 2 box sheared by skewX (0.5): LibreOffice keeps its top edge
        // and moves its bottom edge left, so its centre moves left by
        // tan 0.5 (half the height times the shear) before the translation.
        let (rotation, corner) =
            transform_rotation(Some("skewX (0.5) translate (5 5)"), (0.0, 0.0), 4.0, 2.0);
        let (x, y) = corner.unwrap();
        assert_eq!(rotation, 0.0);
        assert!(
            (x - (5.0 - 0.5f64.tan())).abs() < 1e-9 && (y - 5.0).abs() < 1e-9,
            "{x} {y}"
        );
    }
}
