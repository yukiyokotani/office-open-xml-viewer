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
    by_name: HashMap<(&'static str, String), roxmltree::Node<'a, 'a>>,
}

impl<'a> Definitions<'a> {
    pub fn read(styles: Option<roxmltree::Node<'a, 'a>>) -> Self {
        let mut by_name = HashMap::new();
        if let Some(styles) = styles {
            for node in styles.children().filter(|n| n.is_element()) {
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
                    by_name.entry((kind, name.to_string())).or_insert(node);
                }
            }
        }
        Self { by_name }
    }

    fn get(&self, kind: &'static str, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
        self.by_name.get(&(kind, name.to_string())).copied()
    }
}

/// A resolved graphic style (§16.2 over `style:graphic-properties`).
struct Graphic<'a> {
    cascade: Cascade<'a>,
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
        let mut chain: Vec<StyleRef<'a>> = match node
            .attribute((NS_DRAW, "style-name"))
            .and_then(|name| self.sheet.find(self.scope, Family::Graphic, name))
        {
            Some(style) => self.style_chain(style).to_vec(),
            None => Vec::new(),
        };
        if let Some(default) = self.sheet.default_style(Family::Graphic) {
            chain.push(default);
        }
        self.charge_cascade(&chain, "graphic-properties")?;
        Ok(Graphic {
            cascade: Cascade::new(chain, "graphic-properties"),
        })
    }

    /// A drawing element met in content: a run for the paragraph that
    /// carries it, or `None` when it draws nothing.
    pub(super) fn drawing(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        _host: Option<(u32, &Inline<'a>)>,
    ) -> Result<Option<DocRun>, String> {
        self.enter()?;
        let result = match node.tag_name().name() {
            "frame" => self.frame(node),
            // A hyperlinked drawing (§9.7.1): its frame.
            "a" => match node.children().find(|n| n.is_element()) {
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
    ) -> Placement {
        let anchor = node
            .attribute((NS_TEXT, "anchor-type"))
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
        let z_order = node
            .attribute((NS_DRAW, "z-index"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0)
            .saturating_add(251_658_240);
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
        Placement {
            inline: anchor == "as-char",
            horizontal: (horizontal_frame.into(), horizontal),
            vertical: (vertical_frame.into(), vertical),
            wrap,
            side,
            behind,
            distances,
            z_order,
            allow_overlap: graphic.get(NS_LOEXT, "allow-overlap") != Some("false"),
            polygon: contour_polygon(node),
        }
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

    /// A frame's size (§19.540 svg:width/height, §20.188 fo:min-height,
    /// relative widths §20.341).
    fn frame_size(&self, node: roxmltree::Node<'a, 'a>, graphic: &Graphic<'a>) -> (f64, f64) {
        let get = |ns: &str, name: &str| node.attribute((ns, name)).and_then(|v| length_pt(v).ok());
        let rel = |name: &str| {
            node.attribute((NS_STYLE, name))
                .filter(|v| v.ends_with('%'))
                .and_then(|v| percent(v).ok())
        };
        let width = get(NS_SVG, "width")
            .or_else(|| get(NS_FO, "min-width"))
            .or_else(|| rel("rel-width").map(|r| r * self.text_width))
            .or_else(|| graphic.length(NS_FO, "min-width"))
            .unwrap_or(0.0);
        let height = get(NS_SVG, "height")
            .or_else(|| get(NS_FO, "min-height"))
            .or_else(|| rel("rel-height").map(|r| r * self.text_height))
            .or_else(|| graphic.length(NS_FO, "min-height"))
            .unwrap_or(0.0);
        (width.max(0.0), height.max(0.0))
    }

    /// `draw:frame` (§10.4.2): a picture, a text box or an object's
    /// replacement picture.
    fn frame(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<Option<DocRun>, String> {
        let graphic = self.graphic(node)?;
        let (width, height) = self.frame_size(node, &graphic);
        let placement = self.placement(node, &graphic, (0.0, 0.0));
        let (rotation, _) =
            transform_rotation(node.attribute((NS_DRAW, "transform")), width, height);
        if let Some(text_box) = node.children().find(|n| xml::is(*n, NS_DRAW, "text-box")) {
            return self.text_box(node, text_box, &graphic, placement, width, height, rotation);
        }
        // A formula object (MathML) draws as math in the line.
        if let Some(math) = self.formula_object(node)? {
            return Ok(Some(math));
        }
        // The pictures of the frame: the first renderable one, with an SVG
        // original kept beside its raster replacement.
        let mut raster: Option<(String, &'static str)> = None;
        let mut svg: Option<String> = None;
        for image in node
            .children()
            .filter(|n| xml::is(*n, NS_DRAW, "image"))
            .chain(
                node.children()
                    .filter(|n| {
                        xml::is(*n, NS_DRAW, "object") || xml::is(*n, NS_DRAW, "object-ole")
                    })
                    .flat_map(|object| object.children().filter(|n| xml::is(*n, NS_DRAW, "image"))),
            )
        {
            let resolved = match image.attribute((NS_XLINK, "href")) {
                Some(href) => self.package_image(href)?,
                None => match image
                    .children()
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
        self.charge_construction(std::mem::size_of::<ImageRun>() + path.len() + 64)?;
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
            .children()
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
        let text = self.package.read_text(&content)?;
        self.charge_construction(text.len())?;
        let document = xml::parse(&text)?;
        let root = document.root_element();
        if root.tag_name().name() != "math" {
            return Ok(None);
        }
        // Each referencing frame owns its tree, so each conversion charges
        // every node it builds (the source bytes above cover the parse).
        let nodes = math::convert(root, &mut |bytes| self.charge_construction(bytes))?;
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
        let bytes = &self.images.iter().find(|r| r.key == path)?.bytes;
        let (width, height) = image::intrinsic_size(bytes)?;
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
        let border = graphic
            .get(NS_FO, "border")
            .or_else(|| graphic.get(NS_FO, "border-top"))
            .and_then(|value| parse_border(value, None));
        let border_width = border.as_ref().map_or(0.0, |b| b.width);
        let inner_width = (width - insets[0] - insets[2] - 2.0 * border_width).max(1.0);
        let content = self.contained_story(text_box, inner_width)?;
        let grows = frame.attribute((NS_FO, "min-height")).is_some()
            || text_box.attribute((NS_FO, "min-height")).is_some()
            || graphic.get(NS_DRAW, "auto-grow-height") == Some("true");
        let height = if height <= 0.0 {
            text_box
                .attribute((NS_FO, "min-height"))
                .and_then(|v| length_pt(v).ok())
                .unwrap_or(12.0)
        } else {
            height
        };
        let fill = self.fill(graphic, true);
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
        self.charge_construction(std::mem::size_of::<ShapeRun>() + 128)?;
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
    pub(super) fn background_fill(&mut self, style: &'a crate::style::Style) -> Option<ShapeFill> {
        let graphic = Graphic {
            cascade: Cascade::new(vec![StyleRef::new(style)], "graphic-properties"),
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
    /// companions), or of a frame's `fo:background-color`.
    fn fill(&mut self, graphic: &Graphic<'a>, frame: bool) -> Option<ShapeFill> {
        let color = |value: Option<&str>| value.and_then(|v| hex_color(v).ok());
        match graphic.get(NS_DRAW, "fill") {
            Some("none") => None,
            Some("solid") => {
                color(graphic.get(NS_DRAW, "fill-color")).map(|color| ShapeFill::Solid { color })
            }
            Some("gradient") => graphic
                .get(NS_DRAW, "fill-gradient-name")
                .and_then(|name| self.defs.get("gradient", name))
                .map(gradient),
            Some("hatch") => {
                let hatch = graphic
                    .get(NS_DRAW, "fill-hatch-name")
                    .and_then(|name| self.defs.get("hatch", name));
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
                let href = graphic
                    .get(NS_DRAW, "fill-image-name")
                    .and_then(|name| self.defs.get("fill-image", name))
                    .and_then(|image| image.attribute((NS_XLINK, "href")))?;
                let (path, mime) = self.package_image(href).ok()??;
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
        }
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
        let (rotation, translated) =
            transform_rotation(node.attribute((NS_DRAW, "transform")), width, height);
        if let Some(position) = translated {
            origin = position;
        }
        // Group members carry absolute positions within the group's frame.
        if let Some(offset) = group_offset {
            origin = (origin.0 + offset.0, origin.1 + offset.1);
        }
        let mut placement = self.placement(node, &graphic, (0.0, 0.0));
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
                        x: n(x1, origin.0, width),
                        y: n(y1, origin.1, height),
                    },
                    PathCmd::LineTo {
                        x: n(x2, origin.0, width),
                        y: n(y2, origin.1, height),
                    },
                ]];
            }
            "polyline" | "polygon" => {
                let view = view_box(node).unwrap_or((0.0, 0.0, width.max(1.0), height.max(1.0)));
                let points: Vec<(f64, f64)> = node
                    .attribute((NS_DRAW, "points"))
                    .unwrap_or("")
                    .split_whitespace()
                    .filter_map(|pair| {
                        let (x, y) = pair.split_once(',')?;
                        Some((x.parse().ok()?, y.parse().ok()?))
                    })
                    .collect();
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
                    shape.subpaths = svg_path(d, view);
                }
            }
            "custom-shape" => {
                if let Some(geometry) = node
                    .children()
                    .find(|n| xml::is(*n, NS_DRAW, "enhanced-geometry"))
                {
                    self.enhanced_geometry(geometry, &graphic, &mut shape, width, height);
                } else {
                    shape.preset_geometry = Some("rect".into());
                }
            }
            _ => shape.preset_geometry = Some("rect".into()),
        }
        // Fill and line.
        let closed = !matches!(name, "line" | "connector" | "measure" | "polyline");
        shape.fill = if closed {
            self.fill(&graphic, false)
        } else {
            None
        };
        shape.fill_opacity = graphic
            .get(NS_DRAW, "opacity")
            .and_then(|v| percent(v).ok())
            .filter(|v| *v < 1.0);
        self.line(&graphic, &mut shape);
        // Text in the shape (§10.6.1: paragraphs and lists).
        if node
            .children()
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
            let saved_scope = self.scope;
            let content = self.contained_story(node, inner);
            self.scope = saved_scope;
            shape.text_box_content = content?.into_iter().map(TextBoxBlockWire::Body).collect();
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
        self.charge_construction(
            std::mem::size_of::<ShapeRun>()
                + shape
                    .subpaths
                    .iter()
                    .map(|p| p.len() * std::mem::size_of::<PathCmd>())
                    .sum::<usize>(),
        )?;
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
    ) {
        let kind = node.attribute((NS_DRAW, "type")).unwrap_or("non-primitive");
        shape.flip_h = node.attribute((NS_DRAW, "mirror-horizontal")) == Some("true");
        shape.flip_v = node.attribute((NS_DRAW, "mirror-vertical")) == Some("true");
        let modifiers: Vec<f64> = node
            .attribute((NS_DRAW, "modifiers"))
            .unwrap_or("")
            .split_whitespace()
            .filter_map(|v| v.parse::<f64>().ok())
            .collect();
        if let Some(path) = node.attribute((NS_DRAW, "enhanced-path")) {
            let equations = node
                .children()
                .filter(|n| xml::is(*n, NS_DRAW, "equation"))
                .filter_map(|e| {
                    Some((
                        e.attribute((NS_DRAW, "name"))?.to_string(),
                        e.attribute((NS_DRAW, "formula"))?.to_string(),
                    ))
                })
                .collect();
            let geometry = geometry::Geometry {
                view_box: view_box(node).unwrap_or((0.0, 0.0, 21600.0, 21600.0)),
                equations,
                modifiers,
                log_size: (width * 2540.0 / 72.0, height * 2540.0 / 72.0),
                has_fill: graphic.get(NS_DRAW, "fill") != Some("none"),
                has_stroke: graphic.get(NS_DRAW, "stroke") != Some("none"),
            };
            let path = geometry::evaluate(path, &geometry);
            if !path.subpaths.is_empty() {
                shape.subpaths = path.subpaths;
                if path.paint.iter().any(|p| p.fill.is_some() || !p.stroke) {
                    shape.subpath_paint = path.paint;
                }
                return;
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
    }

    /// The line of a shape (§20.143 draw:stroke and companions).
    fn line(&mut self, graphic: &Graphic<'a>, shape: &mut ShapeRun) {
        let stroke = graphic.get(NS_DRAW, "stroke").unwrap_or("solid");
        if stroke == "none" {
            return;
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
            let dash = graphic
                .get(NS_DRAW, "stroke-dash")
                .and_then(|name| self.defs.get("stroke-dash", name));
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
    }

    /// `draw:g` (§10.3.15): its members as separate shapes on the group's
    /// anchor (DOCX anchors hold one object each).
    fn group(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<Option<DocRun>, String> {
        let mut runs = Vec::new();
        for child in node.children().filter(|n| n.is_element()) {
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
fn contour_polygon(frame: roxmltree::Node<'_, '_>) -> Option<docx_model::AnchorPolygonWire> {
    let contour = frame
        .children()
        .find(|n| xml::is(*n, NS_DRAW, "contour-polygon"))?;
    let view = view_box(contour)?;
    let points: Vec<docx_model::AnchorPointWire> = contour
        .attribute((NS_DRAW, "points"))?
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
    (points.len() >= 3).then_some(docx_model::AnchorPolygonWire {
        edited: false,
        coordinate_space: docx_model::AnchorPolygonSpaceWire {
            width: 21600,
            height: 21600,
        },
        points,
        invalid_point_count: 0,
    })
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

/// `draw:transform` (§19.229): the rotation of `rotate(a) translate(x y)`
/// (radians, counter-clockwise) as DOCX clockwise degrees, and the unrotated
/// box's top-left corner where the transform positions the shape.
fn transform_rotation(value: Option<&str>, width: f64, height: f64) -> (f64, Option<(f64, f64)>) {
    let Some(value) = value else {
        return (0.0, None);
    };
    let mut angle = 0.0;
    let mut translate = None;
    let mut rest = value;
    while let Some(open) = rest.find('(') {
        let name = rest[..open].trim();
        let Some(close) = rest[open..].find(')') else {
            break;
        };
        let arguments: Vec<f64> = rest[open + 1..open + close]
            .split([' ', ','])
            .filter(|v| !v.is_empty())
            .filter_map(|v| length_pt(v).ok().or_else(|| v.parse::<f64>().ok()))
            .collect();
        match name {
            "rotate" => angle += arguments.first().copied().unwrap_or(0.0),
            "translate" => {
                translate = Some((
                    arguments.first().copied().unwrap_or(0.0),
                    arguments.get(1).copied().unwrap_or(0.0),
                ))
            }
            _ => {}
        }
        rest = &rest[open + close + 1..];
    }
    let degrees = (-angle.to_degrees()).rem_euclid(360.0);
    let corner = translate.map(|(tx, ty)| {
        // The centre (w/2, h/2) rotated counter-clockwise by `angle` on the
        // y-down page, then translated.
        let (cx, cy) = (width / 2.0, height / 2.0);
        let (s, c) = angle.sin_cos();
        let center = (cx * c + cy * s + tx, -cx * s + cy * c + ty);
        (center.0 - width / 2.0, center.1 - height / 2.0)
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
        .children()
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

/// An SVG path (`svg:d`, §19.554) in the view box, normalized: absolute and
/// relative M, L, H, V, C, S, Q, T, Z (arcs as lines).
fn svg_path(d: &str, view: (f64, f64, f64, f64)) -> Vec<Vec<PathCmd>> {
    let mut numbers: Vec<(Option<char>, f64)> = Vec::new();
    let mut current = String::new();
    let mut command: Option<char> = None;
    let flush =
        |current: &mut String, command: &mut Option<char>, out: &mut Vec<(Option<char>, f64)>| {
            if let Ok(value) = current.parse::<f64>() {
                out.push((command.take(), value));
            }
            current.clear();
        };
    for c in d.chars() {
        if c.is_ascii_alphabetic() && c != 'e' && c != 'E' {
            flush(&mut current, &mut command, &mut numbers);
            if matches!(c, 'Z' | 'z') {
                numbers.push((Some(c), 0.0));
            } else {
                command = Some(c);
            }
        } else if c == ' ' || c == ',' {
            flush(&mut current, &mut command, &mut numbers);
        } else if c == '-' && !current.is_empty() && !current.ends_with(['e', 'E']) {
            flush(&mut current, &mut command, &mut numbers);
            current.push(c);
        } else {
            current.push(c);
        }
    }
    flush(&mut current, &mut command, &mut numbers);
    let norm = |x: f64, y: f64| ((x - view.0) / view.2, (y - view.1) / view.3);
    let mut paths: Vec<Vec<PathCmd>> = Vec::new();
    let mut path: Vec<PathCmd> = Vec::new();
    let mut point = (0.0, 0.0);
    let mut start = (0.0, 0.0);
    let mut cmd = 'M';
    let mut index = 0;
    while index < numbers.len() {
        if let Some(c) = numbers[index].0 {
            cmd = c;
            if matches!(c, 'Z' | 'z') {
                path.push(PathCmd::Close);
                point = start;
                index += 1;
                continue;
            }
        }
        let relative = cmd.is_ascii_lowercase();
        let arity = match cmd.to_ascii_uppercase() {
            'M' | 'L' | 'T' => 2,
            'H' | 'V' => 1,
            'C' => 6,
            'S' | 'Q' => 4,
            'A' => 7,
            _ => {
                index += 1;
                continue;
            }
        };
        if index + arity > numbers.len() {
            break;
        }
        let v: Vec<f64> = numbers[index..index + arity].iter().map(|n| n.1).collect();
        index += arity;
        let abs = |x: f64, y: f64| {
            if relative {
                (point.0 + x, point.1 + y)
            } else {
                (x, y)
            }
        };
        match cmd.to_ascii_uppercase() {
            'M' => {
                if path.len() > 1 {
                    paths.push(std::mem::take(&mut path));
                } else {
                    path.clear();
                }
                point = abs(v[0], v[1]);
                start = point;
                let (x, y) = norm(point.0, point.1);
                path.push(PathCmd::MoveTo { x, y });
                cmd = if relative { 'l' } else { 'L' };
            }
            'L' | 'T' | 'A' => {
                point = if cmd.eq_ignore_ascii_case(&'A') {
                    abs(v[5], v[6])
                } else {
                    abs(v[0], v[1])
                };
                let (x, y) = norm(point.0, point.1);
                path.push(PathCmd::LineTo { x, y });
            }
            'H' => {
                point.0 = if relative { point.0 + v[0] } else { v[0] };
                let (x, y) = norm(point.0, point.1);
                path.push(PathCmd::LineTo { x, y });
            }
            'V' => {
                point.1 = if relative { point.1 + v[0] } else { v[0] };
                let (x, y) = norm(point.0, point.1);
                path.push(PathCmd::LineTo { x, y });
            }
            'C' => {
                let c1 = abs(v[0], v[1]);
                let c2 = abs(v[2], v[3]);
                let p = abs(v[4], v[5]);
                let ((x1, y1), (x2, y2), (x, y)) =
                    (norm(c1.0, c1.1), norm(c2.0, c2.1), norm(p.0, p.1));
                path.push(PathCmd::CubicBezTo {
                    x1,
                    y1,
                    x2,
                    y2,
                    x,
                    y,
                });
                point = p;
            }
            'S' | 'Q' => {
                let c = abs(v[0], v[1]);
                let p = abs(v[2], v[3]);
                let ((x1, y1), (x, y)) = (norm(c.0, c.1), norm(p.0, p.1));
                path.push(PathCmd::QuadBezTo { x1, y1, x, y });
                point = p;
            }
            _ => {}
        }
    }
    if path.len() > 1 {
        paths.push(path);
    }
    paths
}
