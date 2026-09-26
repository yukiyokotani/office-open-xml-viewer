//! Direct, bounded projection of the validated binary slide tree into the
//! presentation renderer model. No package or XML intermediary is involved.
use super::*;
use ooxml_common::blip::{BlipEffect, SrcRect};
use pptx_model::{
    Fill, PictureElement, ShapeElement, Slide, SlideElement, SlideElementOrigin,
    SlideElementSource, TextBody,
};

const MAX_MODEL_SHAPES: usize = 100_000;
// Bounded upper charge for model-owned paint, identifiers and media path/MIME
// strings. Text and custom-path backing are charged by their own projectors.
const MODEL_SHAPE_STRINGS_BYTES: usize = 384;

/// Whether a text body takes its unspecified properties from the destination
/// main master's `TextMasterStyleAtom` of its own `TextHeaderAtom` type.
///
/// MS-PPT 2.9.44 states this for placeholder shapes. It is silent for a
/// non-placeholder shape whose text type is not `Tx_TYPE_OTHER`, although
/// 2.13.33 names every such type as placeholder text. PowerPoint writes these
/// bodies when a layout-only placeholder is saved to the binary format (with
/// no PlaceholderAtom, or one whose position is -1 per 2.7.8). PowerPoint 16
/// PDF exports of Office-saved decks show these bodies following the master
/// style of their text type, not the document `Tx_TYPE_OTHER` defaults: a
/// type-0 body renders at the master title size and face (44 pt Calibri Light,
/// document default 18 pt Calibri), type-1 bodies render master body bullets
/// absent from the document level, the master body typeface instead of the
/// document one, and the master body's background-scheme text color. No
/// corpus body contradicted the rule. `Tx_TYPE_OTHER` keeps the separately
/// established document fallback below.
fn master_typed_text(text_type: Option<u16>, placeholder: bool) -> bool {
    placeholder || text_type.is_some_and(|kind| kind != 4)
}

fn document_text_axes(
    text_type: Option<u16>,
    placeholder: bool,
    master_linked: bool,
    outline: bool,
    axes: Option<text_style::DocumentAxes>,
) -> Option<text_style::DocumentAxes> {
    (text_type == Some(4) && !placeholder && !master_linked && !outline)
        .then_some(axes)
        .flatten()
}

/// Pattern and texture fills need tile semantics (the pattern bitmap's
/// foreground/background colors, the texture's intrinsic tile size) that the
/// direct model does not derive, and a picture fill with a separate fill
/// rectangle or view-relative placement has no shape-local stretch. Reject
/// them rather than leaving the shape or slide unfilled.
fn unprojected_blip_fill(kind: u32) -> String {
    unsupported(match kind {
        1 => "PowerPoint pattern fill is not projected",
        2 => "PowerPoint texture fill is not projected",
        _ => "PowerPoint picture fill placement is not projected",
    })
}

#[allow(clippy::too_many_arguments)]
pub(in crate::ppt) fn slide(
    index: usize,
    presentation: &persist::OwnedPresentation,
    backing: &[u8],
    pictures: Option<&[u8]>,
    media: &mut media::SpanStore,
    work_budget: &mut usize,
    text_budget: &mut usize,
    model_budget: &mut usize,
) -> Result<Slide, String> {
    let (record, outline) = presentation
        .slides
        .get(index)
        .ok_or_else(|| unsupported("PowerPoint slide index out of range"))?;
    let context = Context {
        presentation,
        backing,
        pictures,
        outline,
        index,
        media,
        work_budget,
        text_budget,
        model_budget,
        elements: Vec::new(),
        sources: Vec::new(),
    };
    context.build(record)
}

struct Context<'a> {
    presentation: &'a persist::OwnedPresentation,
    backing: &'a [u8],
    pictures: Option<&'a [u8]>,
    outline: &'a [String],
    index: usize,
    media: &'a mut media::SpanStore,
    work_budget: &'a mut usize,
    text_budget: &'a mut usize,
    model_budget: &'a mut usize,
    elements: Vec<SlideElement>,
    sources: Vec<SlideElementSource>,
}

impl Context<'_> {
    fn build(mut self, local: &RecordSpan) -> Result<Slide, String> {
        self.media.begin_slide();
        for master in self.presentation.object_masters[self.index].iter() {
            self.drawing(master, true)?;
        }
        self.drawing(local, false)?;
        let scheme = self.presentation.schemes[self.index].as_ref();
        let background = match &self.presentation.backgrounds[self.index] {
            Some(background) => {
                self.charge_shape_strings()?;
                let image = background
                    .paint
                    .background_image()
                    .map(|(id, alpha)| self.image_fill(id, alpha, false))
                    .transpose()?
                    .flatten();
                let gradient = background
                    .paint
                    .project_gradient(
                        &background.gradient.view(self.backing)?,
                        true,
                        scheme,
                        self.work_budget,
                        self.model_budget,
                    )?
                    .map(|value| value.into_model(self.model_budget))
                    .transpose()?;
                if gradient.is_none() && image.is_none() {
                    if let Some(kind) = background.paint.blip_fill_type() {
                        return Err(unprojected_blip_fill(kind));
                    }
                }
                gradient.or_else(|| background.paint.background_model(scheme, image))
            }
            None => None,
        };
        let view = local.view(self.backing)?;
        Ok(Slide {
            index: self.index,
            slide_number: usize::from(self.presentation.first_slide_number) + self.index,
            part_name: None,
            background,
            elements: self.elements,
            element_sources: self.sources,
            notes: None,
            comments: Vec::new(),
            hidden: slide_is_hidden(view.payload, self.work_budget)?,
            parse_error: None,
        })
    }

    fn drawing(&mut self, slide: &RecordSpan, inherited: bool) -> Result<(), String> {
        let mut found = false;
        for record in parse_record_spans(self.backing, slide.payload_span(), self.work_budget)? {
            let view = record.view(self.backing)?;
            if view.kind != 1036 {
                continue;
            }
            if found || view.version != 15 {
                return Err(unsupported("invalid PowerPoint drawing container"));
            }
            found = true;
            for dg in parse_record_spans(self.backing, record.payload_span(), self.work_budget)? {
                let view = dg.view(self.backing)?;
                if view.kind != 0xf002 || view.version != 15 {
                    return Err(unsupported("invalid PowerPoint OfficeArt drawing"));
                }
                for child in parse_record_spans(self.backing, dg.payload_span(), self.work_budget)?
                {
                    self.node(child, false, 0, inherited, &[])?;
                }
            }
        }
        if !found && !inherited {
            return Err(unsupported(
                "PowerPoint direct slide has no drawing; outline fallback is not projected",
            ));
        }
        Ok(())
    }

    fn node(
        &mut self,
        record: RecordSpan,
        nested: bool,
        depth: usize,
        inherited: bool,
        ancestors: &[pptx_model::GroupTransform],
    ) -> Result<(), String> {
        if depth > MAX_DEPTH {
            return Err(unsupported("PowerPoint drawing nesting is too deep"));
        }
        let view = record.view(self.backing)?;
        if view.kind == 0xf003 {
            if view.version != 15 {
                return Err(unsupported("invalid PowerPoint group container"));
            }
            let children =
                parse_record_spans(self.backing, record.payload_span(), self.work_budget)?;
            let first = children
                .first()
                .ok_or_else(|| unsupported("empty PowerPoint group"))?;
            let group = SpannedShape::read_from(
                &SpannedSlideSource {
                    backing: self.backing,
                },
                first.clone(),
                nested,
                self.work_budget,
            )?;
            if group.direct_omitted() || group.props.hidden || (inherited && group.is_placeholder())
            {
                return Ok(());
            }
            if group.is_ole() {
                return Err(unsupported("PowerPoint OLE group shape"));
            }
            let transform = direct_transform::group(&group, nested)?;
            let mut next = ancestors.to_vec();
            if let Some(transform) = transform {
                next.push(transform);
            }
            let child_nested = group.flags & 4 == 0;
            for child in children.into_iter().skip(1) {
                self.node(child, child_nested, depth + 1, inherited, &next)?;
            }
        } else if view.kind == 0xf004 {
            let shape = SpannedShape::read_from(
                &SpannedSlideSource {
                    backing: self.backing,
                },
                record,
                nested,
                self.work_budget,
            )?;
            if shape.direct_omitted() || shape.props.hidden || (inherited && shape.is_placeholder())
            {
                return Ok(());
            }
            self.shape(shape, inherited, ancestors)?;
        }
        Ok(())
    }

    fn shape(
        &mut self,
        shape: SpannedShape,
        inherited: bool,
        ancestors: &[pptx_model::GroupTransform],
    ) -> Result<(), String> {
        if self.elements.len() >= MAX_MODEL_SHAPES {
            return Err(unsupported("too many PowerPoint drawing shapes"));
        }
        let local = direct_transform::leaf(&shape)?;
        let local_extent = (local.cx, local.cy);
        let transform = direct_transform::flatten(local.clone(), ancestors);
        let master = shape
            .master()
            .map(|id| self.presentation.shape_masters.paint(id))
            .transpose()?;
        let paint = master.map_or(shape.props.paint, |base| shape.props.paint.inherit(base));
        let gradient = match shape.master() {
            Some(id) => shape
                .props
                .gradient
                .inherit(self.presentation.shape_masters.gradient(id)?),
            None => shape.props.gradient.clone(),
        }
        .view(self.backing)?;
        let local_geometry = shape.props.geometry.view(self.backing)?;
        let geometry = match shape.master() {
            Some(id) => local_geometry.inherit(
                &self
                    .presentation
                    .shape_masters
                    .geometry(id)?
                    .view(self.backing)?,
            ),
            None => local_geometry,
        };
        let custom = if shape.kind == 75 {
            None
        } else {
            geometry
                .decode(self.work_budget)?
                .as_ref()
                .map(|g| direct_geometry::project(g, self.model_budget))
                .transpose()?
        };
        if shape.is_ole() {
            self.admit_ole_picture(&shape)?;
        }
        if shape.kind == 75 && shape.props.picture_linked {
            return Err(unsupported("linked PowerPoint picture file"));
        }
        if shape.kind == 75 && shape.props.picture != 0 {
            let blip_effects = self.picture_display_effects(&shape, true)?;
            // MS-ODRAW 2.3.23.5: pib names the BLIP displayed by the picture
            // shape. A BLIP this projector cannot carry (PICT, DIB, TIFF, an
            // unused store slot or a mislabeled payload) is rejected rather
            // than leaving an empty frame.
            if self.media.reference(
                shape.props.picture,
                self.backing,
                self.pictures,
                self.work_budget,
            )? {
                self.charge_shape_strings()?;
                let [top, bottom, left, right] = shape.props.crop;
                if left + right >= 100000 || top + bottom >= 100000 {
                    return Err(unsupported("empty PowerPoint picture crop"));
                }
                let (extension, _) = self
                    .media
                    .image(shape.props.picture, self.backing, self.pictures)?
                    .ok_or_else(|| unsupported("PowerPoint picture was not retained"))?;
                let scheme = self.presentation.schemes[self.index].as_ref();
                let (_, stroke) = paint.model_with_custom_geometry(scheme, false, true, None);
                let fill = match paint.picture_backing()? {
                    Some((color, alpha)) => {
                        Some(paint::model_solid(color, alpha, scheme).ok_or_else(|| {
                            unsupported("unresolved PowerPoint picture frame fill color")
                        })?)
                    }
                    None => None,
                };
                self.push(
                    SlideElement::Picture(PictureElement {
                        id: Some(shape.id.to_string()),
                        x: transform.x,
                        y: transform.y,
                        width: transform.cx,
                        height: transform.cy,
                        rotation: transform.rot,
                        flip_h: transform.flip_h,
                        flip_v: transform.flip_v,
                        image_path: format!("legacy-ppt/image/{}", shape.props.picture),
                        mime_type: ooxml_common::blip::mime_from_ext(extension).to_owned(),
                        svg_image_path: None,
                        intrinsic_width_px: None,
                        intrinsic_height_px: None,
                        stroke,
                        fill,
                        prst_geom: None,
                        prst_adjust: None,
                        src_rect: (shape.props.crop != [0; 4]).then_some(SrcRect {
                            l: left as f64 / 100000.0,
                            t: top as f64 / 100000.0,
                            r: right as f64 / 100000.0,
                            b: bottom as f64 / 100000.0,
                        }),
                        alpha: None,
                        duotone: None,
                        blip_effects,
                        cust_geom: None,
                        shadow: None,
                        inner_shadow: None,
                        glow: None,
                        soft_edge: None,
                        reflection: None,
                        scene3d: None,
                        sp3d: None,
                    }),
                    inherited,
                )?;
            } else {
                return Err(unsupported(
                    "PowerPoint picture BLIP is not a supported image",
                ));
            }
        }
        // Alternative shape XML (`ppt::metro`) needs a single blob (two are
        // an ambiguity MS-ODRAW does not resolve) and the master's
        // round-trip theme to resolve it against; a blob without either
        // cannot be verified. Text effects the binary cannot express are
        // deferred: the adopted alternative carries them, and they reject
        // only when it is not adopted.
        let metro_theme = self.presentation.metro_themes[self.index].clone();
        let metro_blob = match (&shape.props.metro, &metro_theme) {
            (None, _) => None,
            (Some(_), _) if shape.props.metro_ambiguous => {
                return Err(crate::ppt::metro::unverifiable("property"));
            }
            (Some(_), None) => return Err(crate::ppt::metro::unverifiable("theme")),
            (Some(span), Some(_)) => Some(span.clone()),
        };
        let deferred_effect = std::cell::Cell::new(None);
        let mut raw_text = None;
        let text = self.text_body(
            &shape,
            inherited,
            &mut raw_text,
            metro_blob.as_ref().map(|_| &deferred_effect),
        )?;
        // Shape types map to presets as PowerPoint converts them; an unmapped
        // type is rejected rather than dropped or drawn as an unfilled box.
        // Adjust values convert on the shape's own (pre-group) extent, where
        // PowerPoint writes its avLst.
        // A picture frame that displayed its picture keeps only its text on
        // an unfilled box, as before.
        let pictured = shape.kind == 75 && shape.props.picture != 0;
        if pictured && text.is_none() {
            return Ok(());
        }
        let mut adjust_bounds = [None; 8];
        let (preset, adjust) = if custom.is_some() || pictured {
            (None, None)
        } else {
            let name = crate::officeart::preset::name(shape.kind).ok_or_else(|| {
                unsupported(format!(
                    "PowerPoint shape type {} has no evidenced preset geometry",
                    shape.kind
                ))
            })?;
            let adjust = crate::officeart::preset::adjustments(
                shape.kind,
                &paint.adjust,
                local_extent.0,
                local_extent.1,
            )?;
            if metro_blob.is_some() {
                // One master unit, the resolution of the anchor the extent
                // was read from (rounded up to whole EMU).
                adjust_bounds = crate::officeart::preset::adjustment_bounds(
                    shape.kind,
                    &paint.adjust,
                    local_extent.0,
                    local_extent.1,
                    1588,
                )?;
            }
            (Some(name), adjust)
        };
        self.charge_shape_strings()?;
        let mut cust_geom_paint = None;
        let fill_area =
            custom.is_some() || (preset.is_some() && !matches!(shape.kind, 20 | 32 | 34 | 38));
        let mut authored_paths = None;
        let (geometry_name, paths, allow_fill, allow_line) = match custom {
            Some(g) => {
                cust_geom_paint = g.paint;
                authored_paths = Some(g.authored);
                ("custGeom".to_owned(), Some(g.paths), g.fill, g.stroke)
            }
            None => (
                preset.unwrap_or("rect").to_owned(),
                None,
                preset.is_some() && !matches!(shape.kind, 20 | 32 | 34 | 38),
                true,
            ),
        };
        let [adj, adj2, adj3, adj4, adj5, adj6, adj7, adj8] = adjust.unwrap_or_default();
        let image = if allow_fill {
            match paint.foreground_image() {
                Some((id, alpha, rotate)) => {
                    self.picture_display_effects(&shape, false)?;
                    self.image_fill(id, alpha, rotate)?
                }
                None => None,
            }
        } else {
            None
        };
        let pattern = match (allow_fill, paint.pattern_image()) {
            (true, Some(pattern)) => {
                // A pattern does not rotate with its shape (fRotateFillWithShape
                // is clear for every projected pattern): PowerPoint's own PPTX
                // save of a 180-degree rotated pattern freeform writes the same
                // tile as for unrotated ones (algn tl, tx/ty 0,
                // rotWithShape 0). Flipped pattern shapes have no evidence.
                if transform.flip_h || transform.flip_v {
                    return Err(unsupported(
                        "PowerPoint pattern fill on a flipped shape is not projected",
                    ));
                }
                Some(self.pattern_fill(pattern)?)
            }
            _ => None,
        };
        if allow_fill && image.is_none() && pattern.is_none() {
            if let Some(kind) = paint.blip_fill_type() {
                return Err(unprojected_blip_fill(kind));
            }
        }
        let (fill, stroke) = paint.model_with_custom_geometry(
            self.presentation.schemes[self.index].as_ref(),
            allow_fill,
            allow_line,
            image,
        );
        let fill = pattern.or(fill);
        // Rotation and flips (own or inherited from groups) do not change the
        // projected shade: in every rotated or flipped corpus shape whose
        // metroBlob shows PowerPoint's own DrawingML (90/180/270 degrees, with
        // and without flips), `ang` stays 90 - fillAngle and only
        // fRotateFillWithShape selects rotWithShape.
        let gradient_fill = paint
            .project_gradient(
                &gradient,
                allow_fill,
                self.presentation.schemes[self.index].as_ref(),
                self.work_budget,
                self.model_budget,
            )?
            .map(|value| value.into_model(self.model_budget))
            .transpose()?;
        let fill = gradient_fill.or(fill);
        let element = ShapeElement {
            x: transform.x,
            y: transform.y,
            width: transform.cx,
            height: transform.cy,
            rotation: transform.rot,
            flip_h: transform.flip_h,
            flip_v: transform.flip_v,
            geometry: geometry_name,
            fill,
            stroke,
            text_body: text,
            default_text_color: None,
            cust_geom: paths,
            cust_geom_paint,
            adj,
            adj2,
            adj3,
            adj4,
            adj5,
            adj6,
            adj7,
            adj8,
            shadow: None,
            inner_shadow: None,
            glow: None,
            soft_edge: None,
            reflection: None,
            id: Some(shape.id.to_string()),
            name: None,
            hyperlink: None,
            hyperlink_action: None,
            placeholder_type: None,
            placeholder_idx: None,
            text_rect: None,
            scene3d: None,
            sp3d: None,
        };
        let adopted = match (&metro_blob, &metro_theme) {
            (Some(span), Some(theme)) => {
                let recorded_fill = if !fill_area {
                    crate::ppt::metro::RecordedFill::NotDisplayed
                } else if !paint.fill_stated() {
                    crate::ppt::metro::RecordedFill::Unstated
                } else if !paint.filled.unwrap_or(true) || !paint.fill_ok.unwrap_or(true) {
                    crate::ppt::metro::RecordedFill::Stated(None)
                } else if allow_fill {
                    crate::ppt::metro::RecordedFill::Stated(element.fill.clone().map(Box::new))
                } else {
                    // Custom geometry whose paths the open-path display rule
                    // leaves unfilled: the recorded fill is still stated.
                    match paint.solid_fill_values(true).and_then(|(color, alpha)| {
                        paint::model_solid(
                            color,
                            alpha,
                            self.presentation.schemes[self.index].as_ref(),
                        )
                    }) {
                        Some(fill) => crate::ppt::metro::RecordedFill::Stated(Some(Box::new(fill))),
                        None => crate::ppt::metro::RecordedFill::Unknown,
                    }
                };
                let binary = crate::ppt::metro::BinaryShape {
                    element: &element,
                    leaf: &local,
                    nested: !ancestors.is_empty(),
                    text: raw_text.as_deref(),
                    fill: recorded_fill,
                    path_paint: authored_paths.take(),
                    adjust_bounds,
                };
                crate::ppt::metro::adopt(
                    &binary,
                    span.view(self.backing)?,
                    theme,
                    self.work_budget,
                    self.text_budget,
                    self.model_budget,
                )?
            }
            _ => None,
        };
        if adopted.is_none() {
            if let Some(name) = deferred_effect.get() {
                return Err(unsupported(format!(
                    "PowerPoint text {name} effect is not projected"
                )));
            }
        }
        self.push(SlideElement::Shape(adopted.unwrap_or(element)), inherited)
    }

    /// OLE shapes (fOleShape, MS-ODRAW 2.2.40) show the presentation picture
    /// the file stores for them: a msosptPictureFrame whose pib names the BLIP
    /// to display (MS-ODRAW 2.3.23.5), resolved through the shape's
    /// ExObjRefAtom (MS-PPT 2.7.7) to the document's ExObjListContainer
    /// (2.10.1). The OLE object itself is never read or activated.
    ///
    /// Evidence and limits: PowerPoint's PDF exports of the local corpus show
    /// the stored picture unchanged for embedded objects with drawAspect
    /// DVASPECT_CONTENT and exColorFollow ExColor_FollowScheme (Excel.Chart
    /// PNG pictures). Icon/thumbnail aspects, linked objects (whose picture may
    /// be refreshed from the link source) and ActiveX controls (drawn live by
    /// the control) are not covered by that evidence and stay rejected.
    fn admit_ole_picture(&self, shape: &SpannedShape) -> Result<(), String> {
        if shape.kind != 75 || shape.props.picture == 0 {
            return Err(unsupported(
                "PowerPoint OLE object has no presentation picture",
            ));
        }
        let id = shape
            .props
            .ole_ref
            .ok_or_else(|| unsupported("PowerPoint OLE shape lacks ExObjRefAtom"))?;
        match self.presentation.ole_objects.get(id)? {
            // [MS-OSHARED] 2.2.1.2 DVASPECT_CONTENT.
            media::OleObject::Embedded { draw_aspect: 1 } => Ok(()),
            media::OleObject::Embedded { .. } => Err(unsupported(
                "PowerPoint OLE object drawn as icon or thumbnail",
            )),
            media::OleObject::Linked => Err(unsupported("linked PowerPoint OLE object")),
            media::OleObject::Control => Err(unsupported("PowerPoint ActiveX control")),
        }
    }

    /// MS-ODRAW 2.3.23 picture display adjustments and MS-PPT 2.7.9 metafile
    /// recoloring, projected onto DrawingML blip effects where PowerPoint's
    /// own reading of the binary is known, otherwise rejected.
    ///
    /// Evidence: PowerPoint 16 opening binary decks (with metroBlobs removed,
    /// so only the binary properties speak) and saving them as PPTX writes
    /// - fPictureGray + fPictureBiLevel as `<a:grayscl/><a:biLevel
    ///   thresh="50000"/>` (PowerPoint's "Black and White");
    /// - pictureTransparent RGB as `<a:clrChange>` from that colour to the
    ///   same colour with alpha 0 (written before the gray/bilevel pair);
    /// - pictureBrightness/pictureContrast as `<a:lum>` (0x8000 -> bright
    ///   100000; 0x599a/0x4ccd -> bright 70000, contrast -70000).
    ///
    /// PowerPoint's PDF exports agree with the first two (black/white by a
    /// luminance threshold, white made transparent). For brightness/contrast,
    /// PowerPoint's PDF of a gray ramp saved as .ppt (with and without its
    /// metroBlobs) renders the same levels as the `<a:lum>` formula of the
    /// presentation renderer. It uses brightness / 0x8000 as the bright
    /// fraction and the 16.16 contrast as the slope k. Contrast is therefore
    /// k - 1 for k <= 1 and 1 - 1/k above. Brightness/contrast together with
    /// the transparent colour or the black-and-white pair has no evidence for
    /// the order PowerPoint applies them in, so that stays rejected. Gray
    /// alone, bi-level alone, recolor and colour modifiers have no evidence
    /// and stay rejected too. Picture fills reject every adjustment.
    fn picture_display_effects(
        &self,
        shape: &SpannedShape,
        picture_frame: bool,
    ) -> Result<Vec<BlipEffect>, String> {
        let p = &shape.props;
        if let Some(adjustment) = p.picture_adjustment {
            return Err(unsupported(format!(
                "PowerPoint picture {adjustment} adjustment is not projected"
            )));
        }
        if p.recolor {
            return Err(unsupported(
                "PowerPoint picture recoloring is not projected",
            ));
        }
        let contrast = p.picture_contrast.filter(|&value| value != 0x10000);
        let brightness = p.picture_brightness.filter(|&value| value != 0);
        let gray = p.picture_gray.unwrap_or(false);
        let bilevel = p.picture_bilevel.unwrap_or(false);
        let transparent = p.picture_transparent;
        let luminance = contrast.is_some() || brightness.is_some();
        if !picture_frame && (gray || bilevel || transparent.is_some() || luminance) {
            return Err(unsupported(
                "PowerPoint picture fill color adjustment is not projected",
            ));
        }
        if luminance {
            if gray || bilevel || transparent.is_some() {
                return Err(unsupported(
                    "PowerPoint picture brightness/contrast combined with other color adjustments is not projected",
                ));
            }
            // MS-ODRAW 2.3.23.11-12: pictureContrast is a 16.16 fixed
            // multiplier (0x10000 = unchanged), pictureBrightness a signed
            // fraction of 0x8000.
            let k = f64::from(contrast.unwrap_or(0x10000)) / 65536.0;
            let contrast = if k <= 1.0 { k - 1.0 } else { 1.0 - 1.0 / k };
            let bright = f64::from(brightness.unwrap_or(0) as i32) / 32768.0;
            return Ok(vec![BlipEffect::Luminance {
                bright: bright.clamp(-1.0, 1.0),
                contrast: contrast.clamp(-1.0, 1.0),
            }]);
        }
        let mut effects = Vec::new();
        if let Some(color) = transparent {
            // OfficeArtCOLORREF (MS-ODRAW 2.2.2): red, green, blue bytes and a
            // flag byte; only a literal RGB colour is a pixel colour to match.
            if color >> 24 != 0 {
                return Err(unsupported(
                    "PowerPoint indexed picture transparent color is not projected",
                ));
            }
            let hex = format!(
                "{:02X}{:02X}{:02X}",
                color & 0xff,
                (color >> 8) & 0xff,
                (color >> 16) & 0xff
            );
            effects.push(BlipEffect::ColorChange {
                from: hex.clone(),
                from_alpha: 1.0,
                to: hex,
                to_alpha: 0.0,
                use_alpha: false,
            });
        }
        match (gray, bilevel) {
            (false, false) => {}
            (true, true) => {
                effects.push(BlipEffect::Grayscale);
                effects.push(BlipEffect::BiLevel { thresh: 0.5 });
            }
            (true, false) => {
                return Err(unsupported(
                    "PowerPoint picture grayscale adjustment is not projected",
                ))
            }
            (false, true) => {
                return Err(unsupported(
                    "PowerPoint picture black-and-white adjustment is not projected",
                ))
            }
        }
        Ok(effects)
    }

    /// msofillPattern as a tiled DrawingML picture fill (ECMA-376 §20.1.8.58).
    ///
    /// Evidence (PowerPoint's PDF export of a deck with pattern fills, and its
    /// own PPTX save of the same binary deck): PowerPoint draws the pattern
    /// BLIP's white pixels in fillColor and its black pixels in fillBackColor
    /// (the PPTX save bakes exactly that mapping into its tiled image), tiles
    /// only the top-left 8x8 pixels of the 10x10 pattern bitmap it stores
    /// (the two transparent rows and columns never appear), draws one pattern
    /// pixel per point, and registers the tiles at the shape's top-left
    /// corner. The black/white mapping is a duotone of the two colours; the
    /// 8x8 area is a source crop; one pixel per point is 72 DPI. Other bitmap
    /// sizes, translucent colours and background patterns have no evidence.
    fn pattern_fill(&mut self, pattern: (u32, u32, u32, u32, u32)) -> Result<Fill, String> {
        let (id, fore, fore_alpha, back, back_alpha) = pattern;
        if fore_alpha != 65536 || back_alpha != 65536 {
            return Err(unsupported(
                "translucent PowerPoint pattern fill is not projected",
            ));
        }
        if !self
            .media
            .reference(id, self.backing, self.pictures, self.work_budget)?
        {
            return Err(unsupported(
                "PowerPoint pattern BLIP is not a supported image",
            ));
        }
        let (extension, bytes) = self
            .media
            .image(id, self.backing, self.pictures)?
            .ok_or_else(|| unsupported("PowerPoint image was not retained"))?;
        let size = match extension {
            "gif" => crate::officeart::raster::gif_size(bytes)?,
            "png" => crate::officeart::raster::png_size(bytes)?,
            _ => {
                return Err(unsupported(
                    "PowerPoint pattern BLIP is not a bitmap pattern",
                ))
            }
        };
        if size != (10, 10) {
            return Err(unsupported(
                "PowerPoint pattern bitmap size is not projected",
            ));
        }
        let scheme = self.presentation.schemes[self.index].as_ref();
        let color = |value| {
            paint::model_color(value, 65536, scheme)
                .ok_or_else(|| unsupported("unresolved PowerPoint pattern color"))
        };
        Ok(Fill::Image {
            image_path: format!("legacy-ppt/image/{id}"),
            mime_type: ooxml_common::blip::mime_from_ext(extension).to_owned(),
            svg_image_path: None,
            dpi: Some(72),
            rot_with_shape: Some(false),
            src_rect: Some(SrcRect {
                l: 0.0,
                t: 0.0,
                r: 0.2,
                b: 0.2,
            }),
            fill_rect: None,
            stretch: false,
            tile: Some(ooxml_common::fill::TileInfo {
                tx: Some(0),
                ty: Some(0),
                sx: Some(1.0),
                sy: Some(1.0),
                flip: Some("none".to_owned()),
                algn: Some("tl".to_owned()),
            }),
            alpha: None,
            // Black -> clr1, white -> clr2 (luminance ramp, §20.1.8.23).
            duotone: Some(ooxml_common::blip::Duotone {
                clr1: color(back)?,
                clr2: color(fore)?,
            }),
            blip_effects: Vec::new(),
        })
    }

    fn image_fill(&mut self, id: u32, opacity: u32, rotate: bool) -> Result<Option<Fill>, String> {
        if !self
            .media
            .reference(id, self.backing, self.pictures, self.work_budget)?
        {
            // MS-ODRAW 2.3.7.6 fillBlip: an unsupported BLIP must not turn a
            // picture fill into an unfilled shape or missing background.
            return Err(unsupported("PowerPoint fill BLIP is not a supported image"));
        }
        let (extension, _) = self
            .media
            .image(id, self.backing, self.pictures)?
            .ok_or_else(|| unsupported("PowerPoint image was not retained"))?;
        Ok(Some(Fill::Image {
            image_path: format!("legacy-ppt/image/{id}"),
            mime_type: ooxml_common::blip::mime_from_ext(extension).to_owned(),
            svg_image_path: None,
            dpi: None,
            rot_with_shape: Some(rotate),
            src_rect: None,
            fill_rect: None,
            stretch: true,
            tile: None,
            alpha: (opacity != 65536).then_some(opacity as f64 / 65536.0),
            duotone: None,
            blip_effects: Vec::new(),
        }))
    }

    fn text_body(
        &mut self,
        shape: &SpannedShape,
        inherited: bool,
        raw_text: &mut Option<String>,
        deferred_effect: Option<&std::cell::Cell<Option<&'static str>>>,
    ) -> Result<Option<TextBody>, String> {
        let Some(textbox) = shape.textbox.as_ref() else {
            return Ok(None);
        };
        let mut blocks = Vec::new();
        let mut style = None;
        let mut text_type = None;
        let mut slide_numbers = Vec::new();
        let mut outline_body = false;
        let mut local_ruler = None;
        let mut ruler_seen = false;
        for atom in parse_record_spans(self.backing, textbox.payload_span(), self.work_budget)? {
            let view = atom.view(self.backing)?;
            match view.kind {
                3999 => {
                    if text_type.is_some() || !blocks.is_empty() {
                        return Err(unsupported("ambiguous PowerPoint text header"));
                    }
                    text_type = Some(text_style::text_type(view)?);
                }
                TEXT_CHARS_ATOM | TEXT_BYTES_ATOM => {
                    if !blocks.is_empty() {
                        return Err(unsupported("duplicate PowerPoint text body"));
                    }
                    let value = decode_text(view)?;
                    charge_text(self.text_budget, value.len())?;
                    blocks.push(value);
                }
                3998 => {
                    if inherited || !blocks.is_empty() || style.is_some() {
                        return Err(unsupported("ambiguous PowerPoint outline text body"));
                    }
                    outline_body = true;
                    let i = u32_at(view.payload, 0)? as usize;
                    text_type = self.presentation.outline_types[self.index].get(i).copied();
                    style = self.presentation.outline_styles[self.index]
                        .get(i)
                        .and_then(|s| s.clone());
                    slide_numbers = self.presentation.outline_slide_numbers[self.index]
                        .get(i)
                        .cloned()
                        .unwrap_or_default();
                    let value = self
                        .outline
                        .get(i)
                        .ok_or_else(|| unsupported("PowerPoint outline text index out of range"))?;
                    charge_text(self.text_budget, value.len())?;
                    blocks.push(value.clone());
                }
                4001 => {
                    if view.version != 0 || style.is_some() || outline_body {
                        return Err(unsupported("invalid PowerPoint text style record"));
                    }
                    style = Some(atom.payload_span().clone());
                }
                4056 => {
                    if blocks.is_empty() || outline_body {
                        return Err(unsupported("orphan PowerPoint slide-number atom"));
                    }
                    slide_numbers.push(text_style::slide_number_position(view)?);
                }
                4006 => {
                    if ruler_seen {
                        return Err(unsupported("duplicate PowerPoint local text ruler"));
                    }
                    ruler_seen = true;
                    local_ruler = Some(ruler::read_full(view, self.work_budget)?);
                }
                _ => {}
            }
        }
        if blocks.is_empty() {
            return Ok(None);
        }
        if blocks.len() != 1 {
            return Err(unsupported("ambiguous PowerPoint styled text body"));
        }
        let master_link = shape.master();
        let linked = master_link
            .map(|id| self.presentation.shape_masters.levels(id))
            .transpose()?;
        let document_axes = document_text_axes(
            text_type,
            shape.is_placeholder(),
            master_link.is_some(),
            outline_body,
            self.presentation.document_text_axes,
        );
        // Master levels resolve through `text_style::master_chain`: typed
        // text through its own, base and document atoms, freeform
        // Tx_TYPE_OTHER text through the document atom alone.
        let master = self.presentation.text_masters[self.index].as_deref();
        let direct = if linked.is_some() {
            None
        } else if master_typed_text(text_type, shape.is_placeholder()) {
            master.and_then(|m| text_type.and_then(|t| m.direct_levels(t)))
        } else if text_type == Some(4) && !shape.is_placeholder() && !outline_body {
            master.and_then(text_style::Master::document_levels)
        } else {
            None
        };
        let levels = linked.or(direct.as_ref().map(|d| d.levels.as_slice()));
        *raw_text = Some(blocks[0].clone());
        slide_numbers.sort_unstable();
        let default_style = style
            .is_none()
            .then(|| text_style::default_style(&blocks[0]));
        let style_bytes = match style.as_ref() {
            Some(span) => span.view(self.backing)?,
            None => default_style
                .as_deref()
                .expect("constructed for absent style"),
        };
        let paragraphs = text_style::direct_model::paragraphs_with_axes(
            &blocks[0],
            style_bytes,
            text_style::Context {
                fonts: &self.presentation.fonts,
                scheme: self.presentation.schemes[self.index].as_ref(),
                levels,
                slide_numbers: &slide_numbers,
                slide_number: u32::from(self.presentation.first_slide_number) + self.index as u32,
                ruler_tabs: local_ruler.and_then(|ruler| ruler.tabs),
                style9: if !outline_body && text_type.is_some() {
                    shape
                        .style9
                        .as_ref()
                        .map(|s| s.view(self.backing))
                        .transpose()?
                } else {
                    None
                },
                deferred_effect,
            },
            text_style::direct_model::DirectAxes {
                ruler: local_ruler,
                document: document_axes,
                ambiguous: direct.as_ref().map(|d| &d.ambiguous),
            },
            self.work_budget,
            self.model_budget,
        )?;
        let p = &shape.props;
        Ok(Some(TextBody {
            vertical_anchor: p.anchor.to_owned(),
            paragraphs,
            default_font_size: None,
            default_bold: None,
            default_italic: None,
            l_ins: i64::from(p.margins[0]),
            r_ins: i64::from(p.margins[2]),
            t_ins: i64::from(p.margins[1]),
            b_ins: i64::from(p.margins[3]),
            wrap: p.wrap.to_owned(),
            // PowerPoint's owned scalar txflTextFlow=1 down-saves both its
            // `vert` and `eaVert` inputs as the same OfficeArt value and
            // round-trips that value as DrawingML eaVert. This compatibility
            // mapping is deliberately independent of text, fonts and outer
            // transforms (MS-ODRAW 2.4.5 uses text-container coordinates).
            // Other flows and nonzero cdirFont were not established by these
            // controls and stay horizontal.
            vert: if p.text_flow == Some(1) && p.font_direction.unwrap_or(0) == 0 {
                "eaVert"
            } else {
                "horz"
            }
            .to_owned(),
            // MS-ODRAW fFitShapeToText is DrawingML spAutoFit (ECMA-376
            // 21.1.2.1.4): the stored anchor is already the fitted size. The
            // classic binary has no shrink-on-overflow (normAutofit) flag.
            auto_fit: if p.fit_shape_to_text { "sp" } else { "none" }.to_owned(),
            font_scale: None,
            ln_spc_reduction: None,
            num_col: 1,
            spc_col: 0,
            rtl_col: false,
            // MS-PPT carries no edge-spacing flag; DrawingML's default
            // (edges suppressed) is the renderer's existing behavior.
            spc_first_last_para: false,
            text_warp: None,
        }))
    }

    fn push(&mut self, element: SlideElement, inherited: bool) -> Result<(), String> {
        let grow = |len: usize, capacity: usize| -> Result<(usize, usize), String> {
            if len < capacity {
                return Ok((capacity, 0));
            }
            let next = capacity
                .max(1)
                .checked_mul(2)
                .ok_or_else(|| unsupported("PowerPoint direct slide model size overflow"))?;
            Ok((next, next - capacity))
        };
        let (element_capacity, element_slots) =
            grow(self.elements.len(), self.elements.capacity())?;
        let (source_capacity, source_slots) = grow(self.sources.len(), self.sources.capacity())?;
        let bytes = element_slots
            .checked_mul(std::mem::size_of::<SlideElement>())
            .and_then(|a| {
                source_slots
                    .checked_mul(std::mem::size_of::<SlideElementSource>())
                    .and_then(|b| a.checked_add(b))
            })
            .ok_or_else(|| unsupported("PowerPoint direct slide model size overflow"))?;
        *self.model_budget = self
            .model_budget
            .checked_sub(bytes)
            .ok_or_else(|| unsupported("PowerPoint direct slide model budget exceeded"))?;
        self.elements
            .try_reserve_exact(element_capacity.saturating_sub(self.elements.len()))
            .map_err(|_| unsupported("PowerPoint direct slide model allocation failed"))?;
        self.sources
            .try_reserve_exact(source_capacity.saturating_sub(self.sources.len()))
            .map_err(|_| unsupported("PowerPoint direct slide model allocation failed"))?;
        self.elements.push(element);
        self.sources.push(SlideElementSource {
            origin: if inherited {
                SlideElementOrigin::Master
            } else {
                SlideElementOrigin::Slide
            },
        });
        Ok(())
    }

    fn charge_shape_strings(&mut self) -> Result<(), String> {
        *self.model_budget = self
            .model_budget
            .checked_sub(MODEL_SHAPE_STRINGS_BYTES)
            .ok_or_else(|| unsupported("PowerPoint direct slide model budget exceeded"))?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(options: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().as_slice(),
            kind.to_le_bytes().as_slice(),
            (payload.len() as u32).to_le_bytes().as_slice(),
            payload,
        ]
        .concat()
    }

    fn properties(values: &[(u16, u32)]) -> Vec<u8> {
        let body: Vec<_> = values
            .iter()
            .flat_map(|(id, value)| id.to_le_bytes().into_iter().chain(value.to_le_bytes()))
            .collect();
        record(((values.len() as u16) << 4) | 3, 0xf00b, &body)
    }

    fn complex_geometry() -> Vec<u8> {
        let array = |size: u16, data: Vec<u8>| {
            let count = (data.len() / usize::from(size)) as u16;
            [
                count.to_le_bytes().as_slice(),
                count.to_le_bytes().as_slice(),
                size.to_le_bytes().as_slice(),
                &data,
            ]
            .concat()
        };
        let vertices = array(
            8,
            [[0i32, 0], [25, 50], [75, 50], [100, 100]]
                .into_iter()
                .flatten()
                .flat_map(i32::to_le_bytes)
                .collect(),
        );
        let segments = array(
            2,
            [0x4000u16, 0x2001, 0x6001, 0x8000]
                .into_iter()
                .flat_map(u16::to_le_bytes)
                .collect(),
        );
        let mut body = Vec::new();
        for (id, value) in [
            (0x140u16, 0),
            (0x141, 0),
            (0x142, 100),
            (0x143, 100),
            (0x8145, vertices.len() as u32),
            (0x8146, segments.len() as u32),
        ] {
            body.extend(id.to_le_bytes());
            body.extend(value.to_le_bytes());
        }
        body.extend(vertices);
        body.extend(segments);
        record((6 << 4) | 3, 0xf00b, &body)
    }

    fn shape(
        kind: u16,
        flags: u32,
        anchor_kind: u16,
        anchor: [i32; 4],
        children: Vec<Vec<u8>>,
    ) -> Vec<u8> {
        let flag = record(
            (kind << 4) | 2,
            0xf00a,
            &[42u32.to_le_bytes(), flags.to_le_bytes()].concat(),
        );
        let anchor = record(
            if anchor_kind == 0xf009 { 1 } else { 0 },
            anchor_kind,
            &anchor
                .into_iter()
                .flat_map(i32::to_le_bytes)
                .collect::<Vec<_>>(),
        );
        record(
            15,
            0xf004,
            &[vec![flag, anchor], children].concat().concat(),
        )
    }

    fn png_blip() -> Vec<u8> {
        let png = vec![
            137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 1,
            8, 6, 0, 0, 0, 244, 34, 127, 138, 0, 0, 0, 14, 73, 68, 65, 84, 120, 156, 99, 248, 207,
            192, 0, 66, 13, 0, 15, 122, 3, 126, 119, 233, 127, 151, 0, 0, 0, 0, 73, 69, 78, 68,
            174, 66, 96, 130,
        ];
        record(0x6e00, 0xf01e, &[vec![0; 17], png].concat())
    }

    fn presentation(span: RecordSpan) -> persist::OwnedPresentation {
        persist::PresentationStorage {
            shape_masters: shape_master::Resolver::default(),
            slides: vec![(span, Vec::new())],
            outline_styles: vec![Vec::new()],
            outline_types: vec![Vec::new()],
            outline_slide_numbers: vec![Vec::new()],
            first_slide_number: 1,
            text_masters: vec![None],
            metro_themes: vec![None],
            document_text_axes: None,
            fonts: Vec::new(),
            schemes: vec![None],
            image_entries: Vec::new(),
            ole_objects: media::OleCatalog::default(),
            backgrounds: vec![None],
            object_masters: vec![std::rc::Rc::from([])],
            size: (720, 540),
        }
    }

    #[test]
    fn document_type4_origins_are_limited_to_unlinked_freeform_text() {
        let mut levels = [text_style::ParagraphAxes::default(); 5];
        levels[0] = text_style::ParagraphAxes {
            margin: Some(180),
            indent: Some(90),
        };
        let axes = Some(levels);
        assert_eq!(document_text_axes(Some(4), false, false, false, axes), axes);
        for excluded in [
            (Some(0), false, false, false),
            (Some(4), true, false, false),
            (Some(4), false, true, false),
            (Some(4), false, false, true),
        ] {
            assert_eq!(
                document_text_axes(excluded.0, excluded.1, excluded.2, excluded.3, axes),
                None
            );
        }
    }

    #[test]
    fn typed_master_style_follows_text_type_not_only_placeholder_status() {
        // Placeholders of any type and non-placeholder text of every
        // placeholder text type use the master style of that type.
        for kind in [0, 1, 2, 5, 6, 7, 8] {
            assert!(master_typed_text(Some(kind), false));
            assert!(master_typed_text(Some(kind), true));
        }
        assert!(master_typed_text(Some(4), true));
        assert!(master_typed_text(None, true));
        // Tx_TYPE_OTHER freeform text keeps the document-default path, and a
        // body without a TextHeaderAtom gains no inferred type.
        assert!(!master_typed_text(Some(4), false));
        assert!(!master_typed_text(None, false));
    }

    #[test]
    fn projects_a_complete_synthetic_slide_without_xml_or_package_round_trip() {
        let flags = record(
            (1 << 4) | 2,
            0xf00a,
            &[42u32.to_le_bytes(), 0u32.to_le_bytes()].concat(),
        );
        let anchor = record(
            0,
            0xf010,
            &[0i16, 0, 576, 576]
                .into_iter()
                .flat_map(i16::to_le_bytes)
                .collect::<Vec<_>>(),
        );
        let style = [
            6u32.to_le_bytes().as_slice(),
            0u16.to_le_bytes().as_slice(),
            0x500u32.to_le_bytes().as_slice(),
            0u16.to_le_bytes().as_slice(),
            0u16.to_le_bytes().as_slice(),
            6u32.to_le_bytes().as_slice(),
            0u32.to_le_bytes().as_slice(),
        ]
        .concat();
        let textbox = record(
            15,
            0xf00d,
            &[
                record(0, TEXT_BYTES_ATOM, b"Hello"),
                record(0, 4001, &style),
            ]
            .concat(),
        );
        let shape = record(
            15,
            0xf004,
            &[
                flags,
                anchor,
                properties(&[(0x181, 0x332211), (0x1c0, 0x665544)]),
                textbox,
            ]
            .concat(),
        );
        let drawing = record(15, 1036, &record(15, 0xf002, &shape));
        let document = record(15, SLIDE_CONTAINER, &drawing);
        let (span, _) = record_span_with_end(&document, 0, &mut 10, "slide").unwrap();
        let presentation = presentation(span);
        let mut media = media::SpanStore::new(Vec::new());
        let model = slide(
            0,
            &presentation,
            &document,
            None,
            &mut media,
            &mut 100,
            &mut 100,
            // The slot charge follows the inline SlideElement size, which
            // tracks the shared chart model; budget one element plus text.
            &mut (2 * std::mem::size_of::<SlideElement>() + 100_000),
        )
        .unwrap();
        assert_eq!(model.slide_number, 1);
        assert_eq!(model.elements.len(), 1);
        assert_eq!(model.element_sources[0].origin, SlideElementOrigin::Slide);
        let SlideElement::Shape(shape) = &model.elements[0] else {
            panic!("shape")
        };
        assert_eq!(
            (shape.x, shape.y, shape.width, shape.height),
            (0, 0, 914400, 914400)
        );
        assert_eq!(shape.geometry, "rect");
        assert!(matches!(shape.fill, Some(Fill::Solid { ref color }) if color == "112233"));
        assert_eq!(shape.stroke.as_ref().unwrap().color, "445566");
        assert_eq!(shape.text_body.as_ref().unwrap().paragraphs.len(), 1);
    }

    #[test]
    fn rejects_model_budget_before_allocating_an_emitted_shape() {
        let shape = record(
            15,
            0xf004,
            &[
                record(
                    (1 << 4) | 2,
                    0xf00a,
                    &[1u32.to_le_bytes(), 0u32.to_le_bytes()].concat(),
                ),
                record(0, 0xf010, &[0u8; 8]),
                properties(&[(0x181, 0)]),
            ]
            .concat(),
        );
        let document = record(
            15,
            SLIDE_CONTAINER,
            &record(15, 1036, &record(15, 0xf002, &shape)),
        );
        let (span, _) = record_span_with_end(&document, 0, &mut 10, "slide").unwrap();
        let presentation = presentation(span);
        let mut media = media::SpanStore::new(Vec::new());
        assert!(slide(
            0,
            &presentation,
            &document,
            None,
            &mut media,
            &mut 100,
            &mut 100,
            &mut 0
        )
        .is_err());
    }

    #[test]
    fn rejects_missing_local_drawing_instead_of_silently_losing_outline_text() {
        let document = record(15, SLIDE_CONTAINER, &[]);
        let (span, _) = record_span_with_end(&document, 0, &mut 10, "slide").unwrap();
        let mut presentation = presentation(span);
        presentation.slides[0].1.push("Visible outline".to_owned());
        let mut media = media::SpanStore::new(Vec::new());
        assert!(slide(
            0,
            &presentation,
            &document,
            None,
            &mut media,
            &mut 100,
            &mut 100,
            &mut 1000
        )
        .unwrap_err()
        .contains("outline fallback"));
    }

    #[test]
    fn integrated_layers_groups_custom_text_and_all_image_consumers_share_resources() {
        let master_shape = shape(
            1,
            0,
            0xf010,
            [0, 0, 100, 100],
            vec![properties(&[(0x181, 0xff)])],
        );
        let master_record = record(
            15,
            SLIDE_CONTAINER,
            &record(15, 1036, &record(15, 0xf002, &master_shape)),
        );
        let style = [
            6u32.to_le_bytes().as_slice(),
            0u16.to_le_bytes().as_slice(),
            0x500u32.to_le_bytes().as_slice(),
            0u16.to_le_bytes().as_slice(),
            0u16.to_le_bytes().as_slice(),
            6u32.to_le_bytes().as_slice(),
            0u32.to_le_bytes().as_slice(),
        ]
        .concat();
        let text = record(
            15,
            0xf00d,
            &[
                record(0, TEXT_BYTES_ATOM, b"Curve"),
                record(0, 4001, &style),
            ]
            .concat(),
        );
        let custom = shape(
            0,
            0,
            0xf00f,
            [10, 20, 110, 120],
            vec![complex_geometry(), properties(&[(0x181, 0x332211)]), text],
        );
        let group_head = shape(
            0,
            1,
            0xf010,
            [0, 0, 1152, 1152],
            vec![record(
                1,
                0xf009,
                &[0i32, 0, 576, 576]
                    .into_iter()
                    .flat_map(i32::to_le_bytes)
                    .collect::<Vec<_>>(),
            )],
        );
        let group = record(15, 0xf003, &[group_head, custom].concat());
        let picture = shape(
            75,
            0x200,
            0xf010,
            [0, 0, 100, 100],
            vec![properties(&[(0x4104, 1)])],
        );
        let image_shape = shape(
            1,
            0,
            0xf010,
            [100, 0, 200, 100],
            vec![properties(&[(0x180, 3), (0x4186, 1), (0x1bf, 0x00100010)])],
        );
        let local_record = record(
            15,
            SLIDE_CONTAINER,
            &record(
                15,
                1036,
                &record(15, 0xf002, &[group, picture, image_shape].concat()),
            ),
        );
        let blip = png_blip();
        let master_offset = 0;
        let local_offset = master_record.len();
        let blip_offset = local_offset + local_record.len();
        let document = [master_record, local_record, blip].concat();
        let master_span = record_span_with_end(&document, master_offset, &mut 20, "master")
            .unwrap()
            .0;
        let local_span = record_span_with_end(&document, local_offset, &mut 20, "slide")
            .unwrap()
            .0;
        let blip_span = record_span_with_end(&document, blip_offset, &mut 20, "blip")
            .unwrap()
            .0;
        let mut p = presentation(local_span);
        p.object_masters[0] = std::rc::Rc::from([master_span]);
        p.image_entries = vec![blip_span];
        let mut background = paint::Paint::default();
        background.property(0x180, 3).unwrap();
        background.property(0x4186, 1).unwrap();
        p.backgrounds[0] = Some(drawing::SpannedBackground {
            paint: background,
            gradient: crate::officeart::gradient::Spanned::default(),
        });
        let mut media = media::SpanStore::new(p.image_entries.clone());
        let model = slide(
            0,
            &p,
            &document,
            None,
            &mut media,
            &mut 1000,
            &mut 1000,
            &mut 1_000_000,
        )
        .unwrap();
        assert_eq!(model.elements.len(), 4);
        assert_eq!(
            model
                .element_sources
                .iter()
                .map(|s| s.origin)
                .collect::<Vec<_>>(),
            vec![
                SlideElementOrigin::Master,
                SlideElementOrigin::Slide,
                SlideElementOrigin::Slide,
                SlideElementOrigin::Slide
            ]
        );
        let SlideElement::Shape(custom) = &model.elements[1] else {
            panic!("custom")
        };
        assert_eq!(custom.geometry, "custGeom");
        assert_eq!(
            (custom.x, custom.y, custom.width, custom.height),
            (31750, 63500, 317500, 317500)
        );
        assert!(matches!(model.elements[2], SlideElement::Picture(_)));
        let SlideElement::Shape(image) = &model.elements[3] else {
            panic!("image fill")
        };
        assert!(
            matches!(image.fill,Some(Fill::Image{ref image_path,..}) if image_path=="legacy-ppt/image/1")
        );
        assert!(
            matches!(model.background,Some(Fill::Image{ref image_path,..}) if image_path=="legacy-ppt/image/1")
        );
        assert_eq!(
            media.used_images().map(|(id, _)| id).collect::<Vec<_>>(),
            vec![1]
        );
        assert_eq!(media.image(1, &document, None).unwrap().unwrap().0, "png");
    }

    /// ExObjListContainer with one OLE container per (container kind,
    /// drawAspect, type, exObjId) tuple (MS-PPT 2.10.1/2.10.12).
    fn ex_obj_list(objects: &[(u16, u32, u32, u32)]) -> Vec<u8> {
        let containers: Vec<u8> = objects
            .iter()
            .flat_map(|&(kind, aspect, object_type, id)| {
                let atom = [aspect, object_type, id, 14, 2, 0]
                    .into_iter()
                    .flat_map(u32::to_le_bytes)
                    .collect::<Vec<_>>();
                record(
                    15,
                    kind,
                    &[
                        record(0, 0x0fcd, &[1, 0, 0, 0, 0, 0, 0, 0]),
                        record(1, 0x0fc3, &atom),
                    ]
                    .concat(),
                )
            })
            .collect();
        record(
            15,
            0x0409,
            &[record(0, 0x040a, &[9, 0, 0, 0]), containers].concat(),
        )
    }

    fn client_data(atoms: &[Vec<u8>]) -> Vec<u8> {
        record(15, 0xf011, &atoms.concat())
    }

    /// One local picture-frame shape on a slide, a one-entry image store and
    /// an optional external object list; returns the projected slide.
    fn project(
        kind: u16,
        flags: u32,
        children: Vec<Vec<u8>>,
        blip: Vec<u8>,
        objects: Option<Vec<u8>>,
    ) -> Result<Slide, String> {
        let picture = shape(kind, flags, 0xf010, [0, 0, 100, 100], children);
        let local = record(
            15,
            SLIDE_CONTAINER,
            &record(15, 1036, &record(15, 0xf002, &picture)),
        );
        let blip_offset = local.len();
        let list_offset = blip_offset + blip.len();
        let document = [local, blip, objects.clone().unwrap_or_default()].concat();
        let local_span = record_span_with_end(&document, 0, &mut 20, "slide")
            .unwrap()
            .0;
        let blip_span = record_span_with_end(&document, blip_offset, &mut 20, "blip")
            .unwrap()
            .0;
        let mut p = presentation(local_span);
        p.image_entries = vec![blip_span];
        if objects.is_some() {
            let list = record_span_with_end(&document, list_offset, &mut 20, "list")
                .unwrap()
                .0;
            p.ole_objects = media::ole_catalog(&document, &[list], &mut 1000);
        }
        let mut media = media::SpanStore::new(p.image_entries.clone());
        slide(
            0,
            &p,
            &document,
            None,
            &mut media,
            &mut 1000,
            &mut 1000,
            &mut 1_000_000,
        )
    }

    const OLE: u32 = 0x200 | 0x10;

    fn ole_ref(id: u32) -> Vec<u8> {
        client_data(&[record(0, 0x0bc1, &id.to_le_bytes())])
    }

    #[test]
    fn embedded_ole_objects_show_their_stored_presentation_picture() {
        let model = project(
            75,
            OLE,
            vec![properties(&[(0x4104, 1), (0x10b, 1)]), ole_ref(7)],
            png_blip(),
            Some(ex_obj_list(&[(0x0fcc, 1, 0, 7)])),
        )
        .unwrap();
        assert_eq!(model.elements.len(), 1);
        let SlideElement::Picture(picture) = &model.elements[0] else {
            panic!("OLE presentation picture")
        };
        assert_eq!(picture.image_path, "legacy-ppt/image/1");
        assert_eq!(picture.mime_type, "image/png");
        assert_eq!((picture.width, picture.height), (158750, 158750));
    }

    #[test]
    fn ole_objects_outside_the_evidenced_display_case_are_rejected() {
        let list = |objects: &[(u16, u32, u32, u32)]| Some(ex_obj_list(objects));
        for (kind, children, objects, expected) in [
            // No picture frame BLIP to show.
            (
                75,
                vec![ole_ref(7)],
                list(&[(0x0fcc, 1, 0, 7)]),
                "no presentation picture",
            ),
            (
                1,
                vec![properties(&[(0x4104, 1)]), ole_ref(7)],
                list(&[(0x0fcc, 1, 0, 7)]),
                "no presentation picture",
            ),
            // Reference resolution (MS-PPT 2.7.7, 2.10.1).
            (
                75,
                vec![properties(&[(0x4104, 1)])],
                list(&[(0x0fcc, 1, 0, 7)]),
                "lacks ExObjRefAtom",
            ),
            (
                75,
                vec![properties(&[(0x4104, 1)]), ole_ref(8)],
                list(&[(0x0fcc, 1, 0, 7)]),
                "unresolved",
            ),
            (
                75,
                vec![properties(&[(0x4104, 1)]), ole_ref(7)],
                None,
                "unresolved",
            ),
            (
                75,
                vec![properties(&[(0x4104, 1)]), ole_ref(7)],
                list(&[(0x0fcc, 1, 0, 7), (0x0fcc, 1, 0, 7)]),
                "ambiguous",
            ),
            (
                75,
                vec![properties(&[(0x4104, 1)]), ole_ref(7)],
                list(&[(0x0fcc, 1, 1, 7)]),
                "inconsistent",
            ),
            // Display cases without Office evidence.
            (
                75,
                vec![properties(&[(0x4104, 1)]), ole_ref(7)],
                list(&[(0x0fcc, 4, 0, 7)]),
                "icon or thumbnail",
            ),
            (
                75,
                vec![properties(&[(0x4104, 1)]), ole_ref(7)],
                list(&[(0x0fce, 1, 1, 7)]),
                "linked",
            ),
            (
                75,
                vec![properties(&[(0x4104, 1)]), ole_ref(7)],
                list(&[(0x0fee, 1, 2, 7)]),
                "ActiveX",
            ),
        ] {
            let error = project(kind, OLE, children, png_blip(), objects).unwrap_err();
            assert!(error.contains(expected), "{expected}: {error}");
        }
        // A malformed list only rejects presentations whose OLE shapes need it.
        let malformed = record(15, 0x0409, &record(15, 0x0fcc, &record(1, 0x0fc3, &[0; 8])));
        assert!(project(
            75,
            0x200,
            vec![properties(&[(0x4104, 1)])],
            png_blip(),
            Some(malformed.clone())
        )
        .is_ok());
        assert!(project(
            75,
            OLE,
            vec![properties(&[(0x4104, 1)]), ole_ref(7)],
            png_blip(),
            Some(malformed)
        )
        .unwrap_err()
        .contains("ExOleObjAtom"));
        // Hidden OLE shapes stay omitted like every other hidden shape.
        let hidden = properties(&[(0x3bf, 0x0002_0002)]);
        assert!(project(75, OLE, vec![hidden], png_blip(), None)
            .unwrap()
            .elements
            .is_empty());
    }

    #[test]
    fn unsupported_or_adjusted_picture_blips_are_rejected_instead_of_dropped() {
        // MS-ODRAW 2.2.23: a DIB BLIP is not carried by the presentation model.
        let dib = record(0x7a80, 0xf01f, &[0; 40]);
        assert!(project(
            75,
            0x200,
            vec![properties(&[(0x4104, 1)])],
            dib.clone(),
            None
        )
        .unwrap_err()
        .contains("picture BLIP"));
        let fill = properties(&[(0x180, 3), (0x4186, 1), (0x1bf, 0x0010_0010)]);
        assert!(project(1, 0x200, vec![fill], dib, None)
            .unwrap_err()
            .contains("fill BLIP"));
        for (values, expected) in [
            (vec![(0x109, 0x8000), (0x13f, 0x0006_0006)], "brightness"),
            (vec![(0x108, 0x4ccd), (0x107, 0x00ff_ffff)], "brightness"),
            (
                vec![(0x107, 0x0800_0001)],
                "indexed picture transparent color",
            ),
            (vec![(0x11a, 0x0000ff)], "recolor"),
            (vec![(0x13f, 0x0004_0004)], "grayscale"),
            (vec![(0x13f, 0x0002_0002)], "black-and-white"),
        ] {
            let values = [vec![(0x4104, 1)], values].concat();
            let error =
                project(75, 0x200, vec![properties(&values)], png_blip(), None).unwrap_err();
            assert!(error.contains(expected), "{expected}: {error}");
        }
        // PowerPoint's reading of the binary: transparent colour, then its
        // "Black and White" pair.
        let values = [(0x4104, 1), (0x107, 0x00c0_8040), (0x13f, 0x0006_0006)];
        let model = project(75, 0x200, vec![properties(&values)], png_blip(), None).unwrap();
        let SlideElement::Picture(picture) = &model.elements[0] else {
            panic!("picture")
        };
        assert_eq!(
            picture.blip_effects,
            vec![
                BlipEffect::ColorChange {
                    from: "4080C0".into(),
                    from_alpha: 1.0,
                    to: "4080C0".into(),
                    to_alpha: 0.0,
                    use_alpha: false,
                },
                BlipEffect::Grayscale,
                BlipEffect::BiLevel { thresh: 0.5 },
            ]
        );
        // Brightness/contrast alone project as <a:lum>: PowerPoint's own
        // reading of 0x599a / 0x4ccd is bright 70000, contrast -70000, and a
        // stored 16.16 slope above 1 maps back to contrast 1 - 1/k.
        for (values, bright, contrast) in [
            (vec![(0x109, 0x599a), (0x108, 0x4ccd)], 0.7, -0.7),
            (vec![(0x109, (-22938i32) as u32)], -0.7, 0.0),
            (vec![(0x108, 218453)], 0.0, 0.7),
        ] {
            let values = [vec![(0x4104, 1)], values].concat();
            let model = project(75, 0x200, vec![properties(&values)], png_blip(), None).unwrap();
            let SlideElement::Picture(picture) = &model.elements[0] else {
                panic!("picture")
            };
            let [BlipEffect::Luminance {
                bright: b,
                contrast: c,
            }] = picture.blip_effects.as_slice()
            else {
                panic!("luminance: {:?}", picture.blip_effects)
            };
            assert!(
                (b - bright).abs() < 1e-3 && (c - contrast).abs() < 1e-3,
                "{b} {c}"
            );
        }
        // Colour adjustments on a picture fill stay rejected.
        let fill = properties(&[
            (0x180, 3),
            (0x4186, 1),
            (0x1bf, 0x0010_0010),
            (0x13f, 0x0006_0006),
        ]);
        assert!(project(1, 0x200, vec![fill], png_blip(), None)
            .unwrap_err()
            .contains("picture fill color"));
        // Contradictory primary/tertiary colour-mode bits are ambiguous.
        let tertiary = {
            let mut bytes = properties(&[(0x13f, 0x0004_0000)]);
            bytes[2..4].copy_from_slice(&0xf122u16.to_le_bytes());
            bytes
        };
        assert!(project(
            75,
            0x200,
            vec![properties(&[(0x4104, 1), (0x13f, 0x0004_0004)]), tertiary],
            png_blip(),
            None
        )
        .unwrap_err()
        .contains("ambiguous"));
        // Default values and unset use bits leave the picture unadjusted.
        for values in [
            vec![
                (0x107, 0xffff_ffff),
                (0x108, 0x10000),
                (0x109, 0),
                (0x117, 0x2000_0000),
            ],
            vec![(0x13f, 0x0006_0000)],
            vec![(0x13f, 0x0000_0006)],
        ] {
            let values = [vec![(0x4104, 1)], values].concat();
            assert_eq!(
                project(75, 0x200, vec![properties(&values)], png_blip(), None)
                    .unwrap()
                    .elements
                    .len(),
                1
            );
        }
        // MS-PPT 2.7.9 RecolorInfoAtom with fShouldRecolor set.
        for (flags, rejected) in [(1u8, true), (0, false)] {
            let atom = record(0, 0x0fe7, &[vec![flags], vec![0; 11]].concat());
            let result = project(
                75,
                0x200,
                vec![properties(&[(0x4104, 1)]), client_data(&[atom])],
                png_blip(),
                None,
            );
            assert_eq!(result.is_err(), rejected);
        }
        // Pattern, texture and non-plain picture fills are not left unfilled.
        for (values, expected) in [
            (vec![(0x180, 1), (0x4186, 1)], Some("pattern bitmap size")),
            (
                vec![(0x180, 1), (0x4186, 1), (0x1bf, 0x0002_0002)],
                Some("pattern fill is not projected"),
            ),
            (vec![(0x180, 2), (0x4186, 1)], Some("texture fill")),
            (
                vec![(0x180, 3), (0x4186, 1), (0x1bf, 0x0002_0002)],
                Some("picture fill placement"),
            ),
            (vec![(0x180, 3)], Some("picture fill placement")),
            (vec![(0x180, 1), (0x4186, 1), (0x1bf, 0x0010_0000)], None),
        ] {
            let result = project(1, 0x200, vec![properties(&values)], png_blip(), None);
            match expected {
                Some(expected) => assert!(result.unwrap_err().contains(expected), "{expected}"),
                None => assert!(matches!(
                    &result.unwrap().elements[0],
                    SlideElement::Shape(shape) if matches!(shape.fill, Some(Fill::None))
                )),
            }
        }
        // Picture-frame backing fill: only an explicit fFilled with a solid
        // fillColor is painted; frames that leave fFilled alone stay unfilled.
        let backing = |values: &[(u16, u32)]| {
            project(
                75,
                0x200,
                vec![properties(&[vec![(0x4104, 1)], values.to_vec()].concat())],
                png_blip(),
                None,
            )
        };
        let picture_fill = |model: Slide| match &model.elements[0] {
            SlideElement::Picture(picture) => picture.fill.clone(),
            _ => panic!("picture"),
        };
        assert!(matches!(
            picture_fill(backing(&[(0x181, 0x4d4d4d), (0x1bf, 0x0010_0010)]).unwrap()),
            Some(Fill::Solid { ref color }) if color == "4D4D4D"
        ));
        for values in [
            vec![(0x181, 0x4d4d4d)],
            vec![(0x181, 0x4d4d4d), (0x1bf, 0x0010_0000)],
            vec![],
        ] {
            assert!(picture_fill(backing(&values).unwrap()).is_none());
        }
        assert!(backing(&[(0x1bf, 0x0010_0010)])
            .unwrap_err()
            .contains("without a color"));
        assert!(backing(&[(0x180, 4), (0x181, 0xff), (0x1bf, 0x0010_0010)])
            .unwrap_err()
            .contains("non-solid"));
        // msofillPattern: a 10x10 pattern bitmap tiles its 8x8 area at one
        // pixel per point, white in fillColor and black in fillBackColor.
        let gif_blip = |w: u16, h: u16| {
            let mut gif = b"GIF89a".to_vec();
            gif.extend(w.to_le_bytes());
            gif.extend(h.to_le_bytes());
            gif.extend([0x80, 0, 0, 0, 0, 0, 0xff, 0xff, 0xff, 0x3b]);
            record(0x6e00, 0xf01e, &[vec![0; 17], gif].concat())
        };
        let pattern = [
            (0x180, 1),
            (0x181, 0x00a87e21),
            (0x183, 0x000d0b0b),
            (0x4186, 1),
        ];
        let model = project(1, 0x200, vec![properties(&pattern)], gif_blip(10, 10), None).unwrap();
        let SlideElement::Shape(shape) = &model.elements[0] else {
            panic!("shape")
        };
        let Some(Fill::Image {
            mime_type,
            dpi,
            src_rect,
            tile,
            duotone,
            stretch,
            ..
        }) = &shape.fill
        else {
            panic!("pattern fill: {:?}", shape.fill)
        };
        assert_eq!(mime_type, "image/gif");
        assert_eq!(*dpi, Some(72));
        assert!(!stretch);
        let crop = src_rect.as_ref().unwrap();
        assert_eq!((crop.l, crop.t, crop.r, crop.b), (0.0, 0.0, 0.2, 0.2));
        let tile = tile.as_ref().unwrap();
        assert_eq!((tile.sx, tile.algn.as_deref()), (Some(1.0), Some("tl")));
        let duotone = duotone.as_ref().unwrap();
        assert_eq!(
            (duotone.clr1.as_str(), duotone.clr2.as_str()),
            ("0B0B0D", "217EA8")
        );
        for (values, blip, flags, expected) in [
            (pattern.to_vec(), gif_blip(8, 8), 0x200, "bitmap size"),
            (
                [pattern.to_vec(), vec![(0x182, 0x8000)]].concat(),
                gif_blip(10, 10),
                0x200,
                "translucent",
            ),
            (pattern.to_vec(), gif_blip(10, 10), 0x240, "flipped shape"),
            (pattern.to_vec(), png_blip(), 0x200, "bitmap size"),
        ] {
            let error = project(1, flags, vec![properties(&values)], blip, None).unwrap_err();
            assert!(error.contains(expected), "{expected}: {error}");
        }
        // A rotated pattern shape keeps the unrotated tile (rotWithShape 0).
        let rotated = [pattern.to_vec(), vec![(0x0004, 180 << 16)]].concat();
        let model = project(1, 0x200, vec![properties(&rotated)], gif_blip(10, 10), None).unwrap();
        let SlideElement::Shape(shape) = &model.elements[0] else {
            panic!("shape")
        };
        assert!(matches!(
            &shape.fill,
            Some(Fill::Image {
                rot_with_shape: Some(false),
                tile: Some(_),
                ..
            })
        ));
        // pib_complex names a linked file rather than a BLIP.
        let mut linked = properties(&[(0xc104, 4)]);
        linked.extend_from_slice(&[b'a', 0, b'b', 0]);
        let len = (linked.len() - 8) as u32;
        linked[4..8].copy_from_slice(&len.to_le_bytes());
        assert!(project(75, 0x200, vec![linked], png_blip(), None)
            .unwrap_err()
            .contains("linked PowerPoint picture"));
    }
}
