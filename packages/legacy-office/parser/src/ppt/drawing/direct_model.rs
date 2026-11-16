//! Direct, bounded projection of the validated binary slide tree into the
//! presentation renderer model. No package or XML intermediary is involved.
use super::*;
use ooxml_common::blip::{BlipEffect, SrcRect};
use pptx_model::{
    Fill, PictureElement, ShapeElement, Slide, SlideElement, SlideElementOrigin,
    SlideElementSource, TextBody, TextRun,
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

/// Ordinary shape texture fills need placement evidence beyond the measured
/// slide-background projection. A picture fill with a separate fill rectangle
/// or view-relative placement has no shape-local stretch. Reject these rather
/// than leaving the shape or slide unfilled.
fn unprojected_blip_fill(kind: u32) -> String {
    unsupported(match kind {
        1 => "PowerPoint pattern fill is not projected",
        2 => "PowerPoint texture fill is not projected",
        _ => "PowerPoint picture fill placement is not projected",
    })
}

/// What a text range does when it is clicked or pointed at ([MS-PPT] 2.6.9
/// InteractiveInfoAtom), as the DrawingML run property that states it.
#[derive(Clone, Copy, Debug, PartialEq)]
enum TextLink<'a> {
    /// `a:hlinkClick` to an external target (action II_HyperlinkAction).
    Click(&'a str),
    /// `a:hlinkClick` with only an action verb (action II_JumpAction).
    Jump(&'static str),
    /// `a:hlinkMouseOver` to an external target.
    Over(&'a str),
}

impl TextLink<'_> {
    fn text(self) -> &'static str {
        match self {
            Self::Jump(action) => action,
            Self::Click(_) | Self::Over(_) => "",
        }
    }

    fn retained_bytes(self) -> usize {
        match self {
            Self::Click(target) | Self::Over(target) => target.len(),
            Self::Jump(action) => action.len(),
        }
    }
}

/// [MS-PPT] 2.6.9 InteractiveInfoAtom.jump (action II_JumpAction) as the
/// `ppaction://hlinkshowjump` verb PowerPoint writes for it (ECMA-376
/// 21.1.2.3.5 `action`). PowerPoint 16 saves jump 2 as
/// `jump=previousslide`; the other verbs follow its DrawingML naming.
fn jump_action(jump: u8) -> Option<&'static str> {
    Some(match jump {
        1 => "ppaction://hlinkshowjump?jump=nextslide",
        2 => "ppaction://hlinkshowjump?jump=previousslide",
        3 => "ppaction://hlinkshowjump?jump=firstslide",
        4 => "ppaction://hlinkshowjump?jump=lastslide",
        5 => "ppaction://hlinkshowjump?jump=lastslideviewed",
        6 => "ppaction://hlinkshowjump?jump=endshow",
        _ => return None,
    })
}

/// [MS-PPT] 2.9.59 ranges are UTF-16 positions in the original text atom;
/// the paragraph separator occupies one position. A range can cut a style
/// run, so split that run at UTF-16-safe character boundaries. The ranges of
/// one call must not overlap: click and mouse-over ranges are applied in
/// separate calls.
fn apply_text_links(
    paragraphs: &mut [pptx_model::Paragraph],
    links: &[(usize, usize, TextLink<'_>)],
    direct_color_authored: &[bool],
    work_budget: &mut usize,
    model_budget: &mut usize,
) -> Result<(), String> {
    // Resource policy: range fan-out and every retained target copy consume
    // the session budgets before allocation. Catalog targets remain borrowed
    // until a projected run needs one. Sorting permits one forward sweep;
    // repeated references never cause a links × runs scan.
    let charge = |budget: &mut usize, bytes: usize, kind: &str| -> Result<(), String> {
        *budget = budget.checked_sub(bytes).ok_or_else(|| {
            crate::rejection::resource(format!("PowerPoint text link {kind} budget exceeded"))
        })?;
        Ok(())
    };
    let scratch = links
        .len()
        .checked_mul(std::mem::size_of::<usize>() + 1)
        .ok_or_else(|| crate::rejection::resource("PowerPoint text link size overflow"))?;
    charge(model_budget, scratch, "model")?;
    charge(
        work_budget,
        links.len().saturating_mul(usize::BITS as usize),
        "work",
    )?;
    let mut order: Vec<_> = (0..links.len()).collect();
    order.sort_unstable_by_key(|&index| links[index].0);
    for pair in order.windows(2) {
        if links[pair[0]].1 > links[pair[1]].0 {
            return Err(unsupported("overlapping PowerPoint text links"));
        }
    }
    let mut seen = vec![false; links.len()];
    let mut next_link = 0usize;
    let mut position = 0usize;
    let count = paragraphs.len();
    for (paragraph_index, paragraph) in paragraphs.iter_mut().enumerate() {
        let mut projected = Vec::new();
        for run in std::mem::take(&mut paragraph.runs) {
            let len = match &run {
                TextRun::Text(data) => data.text.encode_utf16().count(),
                TextRun::Break { .. } => 1,
                TextRun::Math { .. } => return Err(unsupported("PowerPoint linked equation")),
            };
            let end = position
                .checked_add(len)
                .ok_or_else(|| unsupported("PowerPoint text link range overflow"))?;
            if end > direct_color_authored.len() {
                return Err(unsupported("PowerPoint direct text color range"));
            }
            while next_link < order.len() && links[order[next_link]].1 <= position {
                next_link += 1;
            }
            match run {
                TextRun::Text(mut data) if len > 0 => {
                    let active_count = order[next_link..]
                        .iter()
                        .take_while(|&&index| links[index].0 < end)
                        .count();
                    let cut_capacity = len
                        .checked_add(active_count.saturating_mul(2))
                        .and_then(|value| value.checked_add(2))
                        .ok_or_else(|| {
                            crate::rejection::resource("PowerPoint text link size overflow")
                        })?;
                    let scratch = cut_capacity
                        .checked_mul(std::mem::size_of::<usize>())
                        .and_then(|bytes| {
                            (len + 1)
                                .checked_mul(std::mem::size_of::<(usize, usize)>())
                                .and_then(|positions| bytes.checked_add(positions))
                        })
                        .ok_or_else(|| {
                            crate::rejection::resource("PowerPoint text link size overflow")
                        })?;
                    charge(model_budget, scratch, "model")?;
                    let mut cuts = Vec::with_capacity(cut_capacity);
                    cuts.extend([0, len]);
                    for offset in 1..len {
                        if direct_color_authored[position + offset - 1]
                            != direct_color_authored[position + offset]
                        {
                            cuts.push(offset);
                        }
                    }
                    for &index in &order[next_link..next_link + active_count] {
                        for edge in [links[index].0, links[index].1] {
                            if position < edge && edge < end {
                                cuts.push(edge - position);
                            }
                        }
                    }
                    cuts.sort_unstable();
                    cuts.dedup();
                    let source_text = std::mem::take(&mut data.text);
                    let mut positions = Vec::with_capacity(len + 1);
                    positions.push((0, 0));
                    let mut unit = 0;
                    for (byte, ch) in source_text.char_indices() {
                        unit += ch.len_utf16();
                        positions.push((unit, byte + ch.len_utf8()));
                    }
                    // Clone only the style, never the entire original text for
                    // every fragment. JSON size bounds all owned style strings.
                    let style_bytes = usize::try_from(
                        ooxml_common::json_measurement::measure_json(&data)?.json_bytes,
                    )
                    .map_err(|_| {
                        crate::rejection::resource("PowerPoint text link size overflow")
                    })?;
                    let mut selected = next_link;
                    for pair in cuts.windows(2) {
                        charge(work_budget, 1, "work")?;
                        let byte_at = |offset| {
                            positions
                                .binary_search_by_key(&offset, |&(unit, _)| unit)
                                .ok()
                                .map(|index| positions[index].1)
                        };
                        let first = byte_at(pair[0]).ok_or_else(|| {
                            unsupported("PowerPoint text link splits a surrogate")
                        })?;
                        let last = byte_at(pair[1]).ok_or_else(|| {
                            unsupported("PowerPoint text link splits a surrogate")
                        })?;
                        let begin = position + pair[0];
                        let cutoff = position + pair[1];
                        while selected < order.len() && links[order[selected]].1 <= begin {
                            selected += 1;
                        }
                        let link = order
                            .get(selected)
                            .copied()
                            .filter(|&index| links[index].0 <= begin && cutoff <= links[index].1);
                        let target_bytes = link.map_or(0, |index| links[index].2.retained_bytes());
                        let owned_bytes = style_bytes
                            .checked_add(last - first)
                            .and_then(|bytes| bytes.checked_add(target_bytes))
                            .ok_or_else(|| {
                                crate::rejection::resource("PowerPoint text link size overflow")
                            })?;
                        charge(model_budget, owned_bytes, "model")?;
                        text_style::direct_model::reserve_run_slot(&mut projected, model_budget)?;
                        let mut part = data.clone();
                        part.text = source_text[first..last].to_owned();
                        if let Some(index) = link {
                            match links[index].2 {
                                TextLink::Click(target) => {
                                    if part.hyperlink.is_some() || part.hyperlink_action.is_some() {
                                        return Err(unsupported(
                                            "overlapping PowerPoint text links",
                                        ));
                                    }
                                    part.hyperlink = Some(target.to_owned());
                                    // [MS-PPT] 2.9.15 CFMasks.color marks direct color.
                                    // Authored colors survive; inherited fallback uses
                                    // the master's hlink slot (existing Office controls).
                                    if !direct_color_authored[begin] {
                                        part.color = None;
                                    }
                                }
                                link @ TextLink::Jump(_) => {
                                    if part.hyperlink.is_some() || part.hyperlink_action.is_some() {
                                        return Err(unsupported(
                                            "overlapping PowerPoint text links",
                                        ));
                                    }
                                    part.hyperlink_action = Some(link.text().to_owned());
                                }
                                TextLink::Over(target) => {
                                    if part.hyperlink_mouse_over.is_some() {
                                        return Err(unsupported(
                                            "overlapping PowerPoint text links",
                                        ));
                                    }
                                    part.hyperlink_mouse_over = Some(target.to_owned());
                                }
                            }
                            seen[index] = true;
                        }
                        projected.push(TextRun::Text(part));
                    }
                }
                other => {
                    text_style::direct_model::reserve_run_slot(&mut projected, model_budget)?;
                    projected.push(other);
                }
            }
            position = end;
        }
        paragraph.runs = projected;
        if paragraph_index + 1 < count {
            position += 1;
        }
    }
    if seen.iter().any(|seen| !seen) {
        return Err(unsupported("PowerPoint text link range has no run"));
    }
    Ok(())
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
                let image = if let Some((id, alpha)) = background.paint.texture_image() {
                    self.texture_fill(id, alpha, true)?
                } else if let Some(pattern) = background.paint.pattern_image() {
                    Some(self.pattern_fill(pattern, None)?)
                } else {
                    background
                        .paint
                        .background_image()
                        .map(|(id, alpha)| self.image_fill(id, alpha, false))
                        .transpose()?
                        .flatten()
                };
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
                let pattern = image
                    .as_ref()
                    .filter(|_| background.paint.fill_type == Some(1))
                    .cloned();
                gradient
                    .or(pattern)
                    .or_else(|| background.paint.background_model(scheme, image))
            }
            None => None,
        };
        let view = local.view(self.backing)?;
        Ok(Slide {
            theme: None,
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
            return Err(crate::rejection::resource(
                "PowerPoint drawing nesting is too deep",
            ));
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
            return Err(crate::rejection::resource(
                "too many PowerPoint drawing shapes",
            ));
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
        if shape.is_ole() && !self.admit_ole_picture(&shape)? {
            let scheme = self.presentation.schemes[self.index].as_ref();
            let (_, stroke) = paint.model_with_custom_geometry(scheme, false, true, None);
            return self.placeholder_frame(&shape, &transform, stroke, inherited);
        }
        if shape.kind == 75 && shape.props.picture_linked {
            return Err(unsupported("linked PowerPoint picture file"));
        }
        if shape.kind == 75 && shape.props.picture != 0 {
            let blip_effects = self.picture_display_effects(&shape, true)?;
            // MS-ODRAW 2.3.23.5: pib names the BLIP displayed by the picture
            // shape. Opaque PICT is retained for an explicit unavailable-image
            // placeholder; this is recovery policy, not QuickDraw decoding.
            // Other unsupported, absent or mislabeled slots still reject.
            if self.media.reference_picture(
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
                let undecoded = extension == crate::officeart::raster::UNDECODED_JPEG;
                let fill = if undecoded {
                    None
                } else if paint.frame_filled() && paint.fill_type.unwrap_or(0) != 0 {
                    self.picture_frame_fill(&paint, &gradient, &transform)?
                } else {
                    match paint.picture_backing()? {
                        Some((color, alpha)) => {
                            Some(paint::model_solid(color, alpha, scheme).ok_or_else(|| {
                                unsupported("unresolved PowerPoint picture frame fill color")
                            })?)
                        }
                        None => None,
                    }
                };
                // The presentation renderer paints a picture's own fill
                // behind it unless that fill is a picture: a pattern, texture
                // or picture fill is drawn by a rectangle under the picture.
                let (fill, backing) = match fill {
                    Some(image @ Fill::Image { .. }) => (None, Some(image)),
                    fill => (fill, None),
                };
                if undecoded {
                    // ITU-T T.81 arithmetic-coded, lossless, hierarchical and
                    // 12-bit frames have no decoder in the viewers this model
                    // is drawn by. PowerPoint 16 prints the arithmetic and
                    // lossless ones; the picture's frame keeps its place.
                    crate::rejection::approximated("unsupported OfficeArt JPEG frame encoding", ());
                    self.placeholder_frame(&shape, &transform, stroke, inherited)?;
                } else {
                    if let Some(backing) = backing {
                        self.rectangle(None, &transform, backing, None, inherited)?;
                    }
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
                            // The marker is internal to legacy PPT picture frames;
                            // do not broaden image-fill or other format admission.
                            mime_type: if extension == "pict" {
                                "image/x-pict".to_owned()
                            } else {
                                ooxml_common::blip::mime_from_ext(extension).to_owned()
                            },
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
                }
            } else {
                return Err(unsupported(
                    "PowerPoint picture BLIP is not a supported image",
                ));
            }
        }
        // Alternative shape XML (`ppt::metro`) needs a single blob (two are
        // an ambiguity MS-ODRAW does not resolve) and the master's
        // round-trip theme to resolve an actual alternative against. A
        // checksum-only blob needs no theme and retains the binary. Text effects the binary cannot express are
        // deferred: the adopted alternative carries them, and they reject
        // only when it is not adopted.
        let metro_theme = self.presentation.metro_themes[self.index].clone();
        let metro_blob = match (&shape.props.metro, &metro_theme) {
            (None, _) => None,
            (Some(_), _) if shape.props.metro_ambiguous => {
                return Err(crate::ppt::metro::unverifiable("property"));
            }
            (Some(span), _) => Some(span.clone()),
        };
        let deferred_effect = std::cell::Cell::new(None);
        let mut raw_text = None;
        let mut direct_size_authored = None;
        let text = self.text_body(
            &shape,
            inherited,
            &mut raw_text,
            &mut direct_size_authored,
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
        let image = match (allow_fill, image.is_none(), paint.texture_image()) {
            (true, true, Some((id, alpha))) => {
                self.picture_display_effects(&shape, false)?;
                self.texture_fill(id, alpha, paint.rotate_fill_with_shape.unwrap_or(false))?
            }
            _ => image,
        };
        let pattern = match (allow_fill, paint.pattern_image()) {
            // A pattern does not rotate with its shape (fRotateFillWithShape
            // is clear for every projected pattern): PowerPoint's own PPTX
            // save of a 180-degree rotated pattern freeform writes the same
            // tile as for unrotated ones (algn tl, tx/ty 0, rotWithShape 0),
            // and its PDFs of a 30-degree rotated and of a flipped pattern
            // shape print the pattern cells upright and unmirrored
            // (`pattern_fill` writes the bitmap a flipped shape needs).
            (true, Some(pattern)) => Some(self.pattern_fill(pattern, Some(&transform))?),
            _ => None,
        };
        if allow_fill && image.is_none() && pattern.is_none() {
            if let Some(kind) = paint.blip_fill_type() {
                return Err(unprojected_blip_fill(kind));
            }
        }
        let texture = image.clone().filter(|_| paint.fill_type == Some(2));
        let (fill, stroke) = paint.model_with_custom_geometry(
            self.presentation.schemes[self.index].as_ref(),
            allow_fill,
            allow_line,
            image,
        );
        let fill = pattern.or(texture).or(fill);
        // Rotation and flips (own or inherited from groups) do not change the
        // projected shade: in every rotated or flipped corpus shape whose
        // metroBlob shows PowerPoint's own DrawingML (90/180/270 degrees, with
        // and without flips), `ang` stays 90 - fillAngle and only
        // fRotateFillWithShape selects rotWithShape.
        let gradient_fill = paint.project_gradient(
            &gradient,
            allow_fill,
            self.presentation.schemes[self.index].as_ref(),
            self.work_budget,
            self.model_budget,
        )?;
        // A shade edited after PowerPoint wrote it is drawn from the binary
        // (`ResolvedGradient::binary_edited`): its alternative is not read.
        let binary_edited = gradient_fill
            .as_ref()
            .is_some_and(|gradient| gradient.binary_edited());
        let gradient_fill = gradient_fill
            .map(|value| value.into_model(self.model_budget))
            .transpose()?;
        let fill = gradient_fill.or(fill);
        let element = ShapeElement {
            math_fallback: None,
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
        // The picture of a binary picture, pattern or texture fill, for
        // comparison with the alternative's. One that is too long to equal
        // an alternative's, or past the session's comparison allowance, is
        // not compared: the shape keeps its binary projection.
        let image_filled = matches!(element.fill, Some(Fill::Image { .. }));
        let comparison_image = if metro_blob.is_some() && !binary_edited && image_filled {
            let id = paint
                .foreground_image()
                .map(|(id, ..)| id)
                .or_else(|| paint.pattern_image().map(|(id, ..)| id))
                .or_else(|| paint.texture_image().map(|(id, ..)| id));
            id.map(|id| self.media.comparison_copy(id, self.backing, self.pictures))
                .transpose()?
                .flatten()
        } else {
            None
        };
        let uncompared = image_filled && comparison_image.is_none();
        let adopted = match (&metro_blob, &metro_theme) {
            (Some(span), theme) if !binary_edited && !uncompared => {
                let missing_theme = crate::ppt::metro::Theme::Unreadable;
                let theme = theme.as_deref().unwrap_or(&missing_theme);
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
                    direct_size_authored: direct_size_authored.as_deref(),
                    fill: recorded_fill,
                    path_paint: authored_paths.take(),
                    adjust_bounds,
                    image_bytes: comparison_image,
                };
                crate::ppt::metro::adopt(
                    &binary,
                    span.view(self.backing)?,
                    theme,
                    self.media,
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
    /// Evidence: PowerPoint's PDF exports show the stored picture unchanged
    /// for embedded objects with drawAspect DVASPECT_CONTENT (Excel.Chart
    /// PNG pictures) and for a linked object (a linked worksheet range whose
    /// source is not available). An object shown as an icon or thumbnail
    /// stores the picture of that aspect in the same place; PowerPoint could
    /// not export the control, and an ActiveX control could not be authored,
    /// so both are drawn from their stored picture without Office evidence.
    ///
    /// Returns whether the shape has a stored picture to draw; without one
    /// the caller draws a placeholder frame.
    fn admit_ole_picture(&self, shape: &SpannedShape) -> Result<bool, String> {
        let id = shape
            .props
            .ole_ref
            .ok_or_else(|| unsupported("PowerPoint OLE shape lacks ExObjRefAtom"))?;
        match self.presentation.ole_objects.get(id)? {
            // [MS-OSHARED] 2.2.1.2 DVASPECT_CONTENT.
            media::OleObject::Embedded { draw_aspect: 1 } | media::OleObject::Linked => {}
            media::OleObject::Embedded { .. } => crate::rejection::approximated(
                "PowerPoint OLE object drawn as icon or thumbnail",
                (),
            ),
            media::OleObject::Control => {
                crate::rejection::approximated("PowerPoint ActiveX control", ())
            }
        }
        let stored = shape.kind == 75 && shape.props.picture != 0;
        if !stored {
            crate::rejection::approximated("PowerPoint OLE object has no presentation picture", ());
        }
        Ok(stored)
    }

    /// The frame drawn in place of content this reader cannot show: the
    /// shape's own box and line, or a thin gray outline when it has none,
    /// without a fill. The geometry around it is unchanged.
    fn placeholder_frame(
        &mut self,
        shape: &SpannedShape,
        transform: &pptx_model::Transform,
        stroke: Option<pptx_model::Stroke>,
        inherited: bool,
    ) -> Result<(), String> {
        let stroke = stroke.unwrap_or_else(|| pptx_model::Stroke {
            color: "808080".to_owned(),
            width: 9525,
            fill: None,
            dash_style: None,
            custom_dash: Vec::new(),
            line_cap: None,
            line_join: None,
            miter_limit: None,
            alignment: None,
            head_end: None,
            tail_end: None,
            cmpd: None,
        });
        self.rectangle(
            Some(shape.id.to_string()),
            transform,
            Fill::None,
            Some(stroke),
            inherited,
        )
    }

    /// A rectangle over the box of `transform`, without text.
    fn rectangle(
        &mut self,
        id: Option<String>,
        transform: &pptx_model::Transform,
        fill: Fill,
        stroke: Option<pptx_model::Stroke>,
        inherited: bool,
    ) -> Result<(), String> {
        self.charge_shape_strings()?;
        self.push(
            SlideElement::Shape(ShapeElement {
                math_fallback: None,
                x: transform.x,
                y: transform.y,
                width: transform.cx,
                height: transform.cy,
                rotation: transform.rot,
                flip_h: transform.flip_h,
                flip_v: transform.flip_v,
                geometry: "rect".to_owned(),
                fill: Some(fill),
                stroke,
                text_body: None,
                default_text_color: None,
                cust_geom: None,
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
                id,
                name: None,
                hyperlink: None,
                hyperlink_action: None,
                placeholder_type: None,
                placeholder_idx: None,
                text_rect: None,
                scene3d: None,
                sp3d: None,
            }),
            inherited,
        )
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
    /// k - 1 for k <= 1 and 1 - 1/k above. fPictureGray alone prints as a
    /// gray picture, and after brightness/contrast when both are set.
    /// Brightness/contrast together with the transparent colour or the
    /// black-and-white pair has no evidence for the order PowerPoint applies
    /// them in, so that stays rejected. Bi-level alone, recolor and colour
    /// modifiers have no evidence and stay rejected too. Picture fills reject
    /// every adjustment.
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
        let mut effects = Vec::new();
        if luminance {
            if bilevel || transparent.is_some() {
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
            effects.push(BlipEffect::Luminance {
                bright: bright.clamp(-1.0, 1.0),
                contrast: contrast.clamp(-1.0, 1.0),
            });
            if !gray {
                return Ok(effects);
            }
        }
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
            // fPictureGray alone is `<a:grayscl/>`. With brightness the gray
            // conversion comes last: PowerPoint's PDF of a (200, 30, 30)
            // picture with brightness 0x3333 prints gray 158, the luminance
            // of the brightened (255, 132, 132), where brightening its gray
            // 66 would print 168.
            (true, false) => effects.push(BlipEffect::Grayscale),
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
    /// Evidence (PowerPoint's PDF export of decks with pattern fills, and its
    /// own PPTX save of the same binary deck): PowerPoint draws the pattern
    /// BLIP's white pixels in fillColor and its black pixels in fillBackColor
    /// (the PPTX save bakes exactly that mapping into its tiled image) and
    /// registers the tiles at the shape's top-left corner. The black/white
    /// mapping is a duotone of the two colours.
    /// - A 10x10 bitmap (older saves): only its top-left 8x8 pixels tile (the
    ///   two transparent rows and columns never appear), one pattern pixel
    ///   per point. The 8x8 area is a source crop; one pixel per point is
    ///   72 DPI.
    /// - An 8x8 bitmap (PowerPoint 16 saves): the whole bitmap tiles, one
    ///   pattern pixel per 0.75 pt, the size of a pixel at 96 DPI. At 720 DPI
    ///   the PDFs of four controls repeat a four-cell pattern every 30
    ///   pixels.
    ///
    /// fillOpacity and fillBackOpacity apply to their own colour (a control
    /// with both at 50% prints both colours half transparent).
    ///
    /// A bitmap of another size has no evidence: all of it is tiled like an
    /// 8x8 one.
    fn pattern_fill(
        &mut self,
        pattern: (u32, u32, u32, u32, u32),
        frame: Option<&pptx_model::Transform>,
    ) -> Result<Fill, String> {
        let (id, fore, fore_alpha, back, back_alpha) = pattern;
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
        let (dpi, src_rect) = match size {
            (8, 8) => (96, None),
            (10, 10) => (
                72,
                Some(SrcRect {
                    l: 0.0,
                    t: 0.0,
                    r: 0.2,
                    b: 0.2,
                }),
            ),
            _ => crate::rejection::approximated(
                "PowerPoint pattern bitmap size is not projected",
                (96, None),
            ),
        };
        // The presentation renderer reflects a picture fill with its shape.
        // PowerPoint does not reflect a pattern (see the caller), so a
        // flipped shape tiles the bitmap as seen through its own reflections,
        // from the corner that is drawn at the top left: the tile grid is
        // offset by the shape's extent modulo the 8-pixel tile.
        let (flip_h, flip_v) = frame.map_or((false, false), |frame| (frame.flip_h, frame.flip_v));
        let mut reflected = None;
        if flip_h || flip_v {
            *self.work_budget = self.work_budget.checked_sub(64).ok_or_else(|| {
                crate::rejection::resource("PowerPoint pattern work budget exceeded")
            })?;
            reflected = self.media.reflected_pattern(
                id,
                flip_h,
                flip_v,
                |extension, bytes| {
                    Ok(crate::ppt::pattern::cells(extension, bytes)?.map(|cells| {
                        crate::ppt::pattern::png(crate::ppt::pattern::mirrored(
                            cells, flip_h, flip_v,
                        ))
                    }))
                },
                self.backing,
                self.pictures,
            )?;
            if reflected.is_none() {
                // Not a bitmap this reader can rewrite: it tiles reflected.
                crate::rejection::approximated(
                    "PowerPoint pattern fill on a flipped shape is not projected",
                    (),
                );
            }
        }
        let (id, extension, src_rect) = match reflected {
            Some(id) => (id, "png", None),
            None => (id, extension, src_rect),
        };
        let tile_emu = 8 * 914_400 / i64::from(dpi);
        let offset = |flipped: bool, extent: i64| {
            if flipped && reflected.is_some() {
                extent.rem_euclid(tile_emu)
            } else {
                0
            }
        };
        let (tx, ty) = frame.map_or((0, 0), |frame| {
            (offset(flip_h, frame.cx), offset(flip_v, frame.cy))
        });
        let scheme = self.presentation.schemes[self.index].as_ref();
        let color = |value, alpha| {
            paint::model_color(value, alpha, scheme)
                .ok_or_else(|| unsupported("unresolved PowerPoint pattern color"))
        };
        Ok(Fill::Image {
            image_path: format!("legacy-ppt/image/{id}"),
            mime_type: ooxml_common::blip::mime_from_ext(extension).to_owned(),
            svg_image_path: None,
            dpi: Some(dpi),
            rot_with_shape: Some(false),
            src_rect,
            fill_rect: None,
            stretch: false,
            tile: Some(ooxml_common::fill::TileInfo {
                tx: Some(tx),
                ty: Some(ty),
                sx: Some(1.0),
                sy: Some(1.0),
                flip: Some("none".to_owned()),
                algn: Some("tl".to_owned()),
            }),
            alpha: None,
            // Black -> clr1, white -> clr2 (luminance ramp, §20.1.8.23).
            duotone: Some(ooxml_common::blip::Duotone {
                clr1: color(back, back_alpha)?,
                clr2: color(fore, fore_alpha)?,
            }),
            blip_effects: Vec::new(),
        })
    }

    /// msofillTexture (MS-ODRAW 2.4.11) as a tiled DrawingML picture fill.
    ///
    /// PowerPoint 16 saves a texture as a 100% intrinsic tile at (0,0)
    /// (`<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none"
    /// algn="tl"/>`), for slide backgrounds even when the binary origin and
    /// EMU tile dimensions were changed, and for shapes with a preset or a
    /// user texture; its PDF exports tile them so.
    fn texture_fill(
        &mut self,
        id: u32,
        opacity: u32,
        rotate: bool,
    ) -> Result<Option<Fill>, String> {
        let mut image = self.image_fill(id, opacity, rotate)?;
        if let Some(Fill::Image {
            dpi, stretch, tile, ..
        }) = image.as_mut()
        {
            *dpi = Some(0);
            *stretch = false;
            *tile = Some(ooxml_common::fill::TileInfo {
                tx: Some(0),
                ty: Some(0),
                sx: Some(1.0),
                sy: Some(1.0),
                flip: Some("none".to_owned()),
                algn: Some("tl".to_owned()),
            });
        }
        Ok(image)
    }

    /// The fill of a picture frame that is not a solid colour. PowerPoint
    /// paints a frame's fill behind the transparent pixels of its picture
    /// (`Paint::picture_backing`); its PDFs of frames with a shade or a
    /// pattern behind an opaque picture show the picture only, so the fill
    /// kinds are projected as for a shape, without Office evidence of their
    /// placement behind a picture. The caller draws a picture-type fill with
    /// a rectangle under the picture.
    fn picture_frame_fill(
        &mut self,
        paint: &paint::Paint,
        gradient: &crate::officeart::gradient::Borrowed<'_>,
        frame: &pptx_model::Transform,
    ) -> Result<Option<Fill>, String> {
        const SITE: &str = "PowerPoint picture frame non-solid fill";
        let scheme = self.presentation.schemes[self.index].as_ref();
        let fill = if let Some(shade) =
            paint.project_gradient(gradient, true, scheme, self.work_budget, self.model_budget)?
        {
            Some(shade.into_model(self.model_budget)?)
        } else if let Some(pattern) = paint.pattern_image() {
            Some(self.pattern_fill(pattern, Some(frame))?)
        } else if let Some((id, alpha)) = paint.texture_image() {
            self.texture_fill(id, alpha, paint.rotate_fill_with_shape.unwrap_or(false))?
        } else if let Some((id, alpha, rotate)) = paint.foreground_image() {
            self.image_fill(id, alpha, rotate)?
        } else {
            None
        };
        Ok(crate::rejection::approximated(SITE, fill))
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
        direct_size_authored: &mut Option<Vec<bool>>,
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
        let mut pending_action = None;
        let mut text_links = Vec::new();
        let mut range_end = None;
        for atom in parse_record_spans(self.backing, textbox.payload_span(), self.work_budget)? {
            let view = atom.view(self.backing)?;
            match view.kind {
                0x0ff2 if shape.props.metro.is_some() => {
                    // [MS-PPT] 2.6.8-10: a click (instance 0, 2.9.56) or
                    // mouse-over (instance 1, 2.9.58) action immediately
                    // precedes its text-range atom of the same instance
                    // (2.9.57, 2.9.59) in a client textbox.
                    if view.version != 15 || view.instance > 1 || pending_action.is_some() {
                        return Err(unsupported("invalid PowerPoint text interaction"));
                    }
                    let children =
                        parse_record_spans(self.backing, atom.payload_span(), self.work_budget)?;
                    if children.len() != 1 {
                        return Err(unsupported("invalid PowerPoint text interaction"));
                    }
                    let action = children[0].view(self.backing)?;
                    if action.kind != 0x0ff3 || action.version != 0 || action.payload.len() != 16 {
                        return Err(unsupported("invalid PowerPoint text interaction atom"));
                    }
                    pending_action = Some((
                        view.instance,
                        u32_at(action.payload, 4)?,
                        action.payload[8],
                        action.payload[10],
                    ));
                }
                0x0fdf if shape.props.metro.is_some() => {
                    if view.version != 0 || view.instance > 1 || view.payload.len() != 8 {
                        return Err(unsupported("invalid PowerPoint text link range"));
                    }
                    let (instance, id, action, jump) = pending_action
                        .take()
                        .ok_or_else(|| unsupported("orphan PowerPoint text link range"))?;
                    if instance != view.instance {
                        return Err(unsupported("invalid PowerPoint text link range"));
                    }
                    let over = instance == 1;
                    // InteractiveInfoAtom.action: 4 is II_HyperlinkAction and
                    // 3 II_JumpAction. PowerPoint 16 PDF controls draw the
                    // range of a mouse-over hyperlink and of a click jump
                    // like a click hyperlink; a mouse-over jump has no
                    // control and no DrawingML target here, so it states
                    // nothing, like the other actions.
                    let link = match (action, over) {
                        (4, false) => Some(TextLink::Click(self.presentation.hyperlinks.get(id)?)),
                        (4, true) => Some(TextLink::Over(self.presentation.hyperlinks.get(id)?)),
                        (3, false) => jump_action(jump).map(TextLink::Jump),
                        _ => None,
                    };
                    let begin = u32_at(view.payload, 0)? as usize;
                    let end = u32_at(view.payload, 4)? as usize;
                    // [MS-PPT] 2.9.59 requires a range valid for the text,
                    // for every action: its length is end - begin, and its
                    // end is checked against the text once that is known.
                    if begin > end {
                        return Err(crate::rejection::corrupt(
                            "PowerPoint text link range ends before it begins",
                        ));
                    }
                    range_end = Some(range_end.map_or(end, |seen: usize| seen.max(end)));
                    if let Some(target) = link {
                        if begin >= end {
                            return Err(unsupported("invalid PowerPoint text link range"));
                        }
                        *self.model_budget = self
                            .model_budget
                            .checked_sub(2 * std::mem::size_of::<(usize, usize, TextLink<'_>)>())
                            .ok_or_else(|| {
                                crate::rejection::resource(
                                    "PowerPoint text link model budget exceeded",
                                )
                            })?;
                        text_links.push((begin, end, target));
                    }
                }
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
        if let Some(end) = range_end {
            let units = blocks.first().map(|text| text.encode_utf16().count());
            if units.is_none_or(|units| end > units) {
                return Err(crate::rejection::corrupt(
                    "PowerPoint text link range exceeds its text",
                ));
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
        let direct_color_authored = if shape.props.metro.is_some() || !text_links.is_empty() {
            let masks =
                text_style::direct_character_masks(&blocks[0], style_bytes, self.work_budget)?;
            if shape.props.metro.is_some() {
                *direct_size_authored = Some(masks.size);
            }
            Some(masks.color)
        } else {
            None
        };
        let mut paragraphs = text_style::direct_model::paragraphs_with_axes(
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
        if !text_links.is_empty() {
            let colors = direct_color_authored
                .as_deref()
                .ok_or_else(|| unsupported("PowerPoint direct text color mask missing"))?;
            // Click and mouse-over ranges are independent run properties
            // and may cover the same characters.
            for over in [false, true] {
                let links: Vec<_> = text_links
                    .iter()
                    .copied()
                    .filter(|(_, _, link)| matches!(link, TextLink::Over(_)) == over)
                    .collect();
                if !links.is_empty() {
                    apply_text_links(
                        &mut paragraphs,
                        &links,
                        colors,
                        self.work_budget,
                        self.model_budget,
                    )?;
                }
            }
        }
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
            // These DrawingML body properties are not supplied by this
            // binary projector. Alternative shape XML keeps authored values.
            anchor_ctr: false,
            compat_ln_spc: None,
            text_warp: None,
            shrink_on_overflow: false,
        }))
    }

    fn push(&mut self, element: SlideElement, inherited: bool) -> Result<(), String> {
        let grow = |len: usize, capacity: usize| -> Result<(usize, usize), String> {
            if len < capacity {
                return Ok((capacity, 0));
            }
            let next = capacity.max(1).checked_mul(2).ok_or_else(|| {
                crate::rejection::resource("PowerPoint direct slide model size overflow")
            })?;
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
            .ok_or_else(|| {
                crate::rejection::resource("PowerPoint direct slide model size overflow")
            })?;
        *self.model_budget = self.model_budget.checked_sub(bytes).ok_or_else(|| {
            crate::rejection::resource("PowerPoint direct slide model budget exceeded")
        })?;
        self.elements
            .try_reserve_exact(element_capacity.saturating_sub(self.elements.len()))
            .map_err(|_| {
                crate::rejection::resource("PowerPoint direct slide model allocation failed")
            })?;
        self.sources
            .try_reserve_exact(source_capacity.saturating_sub(self.sources.len()))
            .map_err(|_| {
                crate::rejection::resource("PowerPoint direct slide model allocation failed")
            })?;
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
            .ok_or_else(|| {
                crate::rejection::resource("PowerPoint direct slide model budget exceeded")
            })?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn repeated_link_targets_cannot_exceed_the_projection_budget() {
        let mut paragraphs: Vec<pptx_model::Paragraph> =
            serde_json::from_value(serde_json::json!([{
                "alignment":"l", "marL":0, "marR":0, "indent":0,
                "spaceBefore":null, "spaceAfter":null, "spaceLine":null,
                "lvl":0, "bullet":{"type":"none"}, "defFontSize":null,
                "defColor":null, "defBold":null, "defItalic":null,
                "defFontFamily":null, "tabStops":[], "rtl":false, "eaLnBrk":true,
                "runs":[{"type":"text", "text":"abcd", "bold":null, "italic":null,
                    "underline":false, "strikethrough":false, "strikeDouble":false,
                    "fontSize":18.0, "color":null, "fontFamily":null,
                    "fieldType":null, "hyperlinkUsesTextFill":false}]
            }]))
            .unwrap();
        let target = "x".repeat(4096);
        let links = vec![
            (0, 1, TextLink::Click(target.as_str())),
            (1, 2, TextLink::Click(target.as_str())),
            (2, 4, TextLink::Click(target.as_str())),
        ];
        let error = apply_text_links(&mut paragraphs, &links, &[false; 4], &mut 10_000, &mut 4096)
            .unwrap_err();
        assert!(error.contains("model budget"), "{error}");
    }

    #[test]
    fn mouse_over_and_jump_ranges_state_their_own_run_properties() {
        let mut paragraphs: Vec<pptx_model::Paragraph> =
            vec![serde_json::from_value(serde_json::json!({
                "alignment": "l", "marL": 0, "marR": 0, "indent": 0,
                "spaceBefore": null, "spaceAfter": null, "spaceLine": null,
                "lvl": 0, "bullet": {"type": "none"}, "defFontSize": null,
                "defColor": null, "defBold": null, "defItalic": null,
                "defFontFamily": null, "tabStops": [], "rtl": false, "eaLnBrk": true,
                "runs": [{"type": "text", "text": "abcd", "bold": null, "italic": null,
                    "underline": false, "strikethrough": false, "strikeDouble": false,
                    "fontSize": 18.0, "color": "000000", "fontFamily": null,
                    "fieldType": null, "hyperlinkUsesTextFill": false}]
            }))
            .unwrap()];
        let colors = [false; 5];
        // A click jump on "ab", then a mouse-over target on "bc": the two
        // properties overlap on "b" without conflict.
        let jump = jump_action(2).unwrap();
        assert_eq!(jump, "ppaction://hlinkshowjump?jump=previousslide");
        assert_eq!(jump_action(0), None);
        assert_eq!(jump_action(7), None);
        for links in [
            [(0, 2, TextLink::Jump(jump))],
            [(1, 3, TextLink::Over("https://example.invalid/o"))],
        ] {
            apply_text_links(&mut paragraphs, &links, &colors, &mut 10_000, &mut 100_000).unwrap();
        }
        let runs: Vec<_> = paragraphs[0]
            .runs
            .iter()
            .map(|run| match run {
                TextRun::Text(text) => (
                    text.text.as_str(),
                    text.hyperlink.as_deref(),
                    text.hyperlink_action.as_deref(),
                    text.hyperlink_mouse_over.as_deref(),
                    text.color.as_deref(),
                ),
                _ => panic!("unexpected break"),
            })
            .collect();
        let over = Some("https://example.invalid/o");
        // Neither is a click target: the run keeps its colour.
        assert_eq!(
            runs,
            [
                ("a", None, Some(jump), None, Some("000000")),
                ("b", None, Some(jump), over, Some("000000")),
                ("c", None, None, over, Some("000000")),
                ("d", None, None, None, Some("000000")),
            ]
        );
        // Two click-class ranges on the same characters stay a conflict.
        let error = apply_text_links(
            &mut paragraphs,
            &[(0, 1, TextLink::Click("https://example.invalid/"))],
            &colors,
            &mut 10_000,
            &mut 100_000,
        )
        .unwrap_err();
        assert!(error.contains("overlapping"), "{error}");
    }

    #[test]
    fn links_clear_inherited_color_but_keep_explicit_color() {
        let run = serde_json::json!({
            "type": "text", "text": "abcd", "bold": null, "italic": null,
            "underline": false, "underlineStyle": null, "underlineColor": null,
            "strikethrough": false, "strikeDouble": false, "fontSize": 18.0,
            "color": "000000", "fontFamily": "Arial", "fontFamilyEa": null,
            "fontFamilySym": null, "baseline": null, "caps": null,
            "letterSpacing": null, "fieldType": null, "hyperlink": null,
            "hyperlinkUsesTextFill": false, "hyperlinkAction": null, "shadow": null, "reflection": null,
            "outline": null, "highlight": null
        });
        let mut paragraphs: Vec<pptx_model::Paragraph> =
            vec![serde_json::from_value(serde_json::json!({
                "alignment": "l", "marL": 0, "marR": 0, "indent": 0,
                "spaceBefore": null, "spaceAfter": null, "spaceLine": null,
                "lvl": 0, "bullet": {"type": "none"}, "defFontSize": null,
                "defColor": null, "defBold": null, "defItalic": null,
                "defFontFamily": null, "tabStops": [], "rtl": false,
                "eaLnBrk": true, "runs": [run]
            }))
            .unwrap()];
        apply_text_links(
            &mut paragraphs,
            &[(1, 3, TextLink::Click("https://example.invalid"))],
            &[false, false, true, false, false],
            &mut 10_000,
            &mut 100_000,
        )
        .unwrap();
        let text_runs: Vec<_> = paragraphs[0]
            .runs
            .iter()
            .map(|run| match run {
                TextRun::Text(text) => text,
                _ => panic!("unexpected break"),
            })
            .collect();
        assert_eq!(
            text_runs
                .iter()
                .map(|run| run.text.as_str())
                .collect::<Vec<_>>(),
            ["a", "b", "c", "d"]
        );
        assert_eq!(text_runs[0].color.as_deref(), Some("000000"));
        assert_eq!(text_runs[1].color, None);
        assert_eq!(text_runs[2].color.as_deref(), Some("000000"));
        assert_eq!(text_runs[3].color.as_deref(), Some("000000"));
        assert!(text_runs[1].hyperlink.is_some());
        assert!(text_runs[2].hyperlink.is_some());
        assert!(text_runs[0].hyperlink.is_none());
    }

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
            hyperlinks: media::HyperlinkCatalog::default(),
            backgrounds: vec![None],
            object_masters: vec![std::rc::Rc::from([])],
            size: (720, 540),
        }
    }

    #[test]
    fn checksum_only_metro_blob_keeps_binary_shape_without_a_theme() {
        use std::io::Write;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        zip.start_file("_rels/.rels", zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="downrev" Type="http://schemas.microsoft.com/office/2006/relationships/downRev" Target="downrev.xml"/></Relationships>"#).unwrap();
        let blob = zip.finish().unwrap().into_inner();
        let property = record(
            0x13,
            0xf00b,
            &[
                0x83a9u16.to_le_bytes().as_slice(),
                &(blob.len() as u32).to_le_bytes(),
                &blob,
            ]
            .concat(),
        );
        let drawing = record(
            15,
            1036,
            &record(
                15,
                0xf002,
                &shape(1, 0, 0xf010, [0, 0, 100, 100], vec![property]),
            ),
        );
        let document = record(15, SLIDE_CONTAINER, &drawing);
        let (span, _) = record_span_with_end(&document, 0, &mut 100, "slide").unwrap();
        let presentation = presentation(span);
        let result = slide(
            0,
            &presentation,
            &document,
            None,
            &mut media::SpanStore::new(Vec::new()),
            &mut 10_000,
            &mut 100_000,
            &mut (16 * 1024 * 1024),
        )
        .unwrap();
        assert!(
            matches!(result.elements.as_slice(), [SlideElement::Shape(shape)] if shape.geometry == "rect")
        );
    }

    /// A text box with one interaction (MS-PPT 2.6.8-10) on its first two
    /// characters, in a shape that carries a checksum-only metroBlob.
    fn interactive_text(
        instance: u16,
        action: u8,
        jump: u8,
        range_instance: u16,
        range: (u8, u8),
    ) -> Result<Slide, String> {
        use std::io::Write;
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        zip.start_file("_rels/.rels", zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="downrev" Type="http://schemas.microsoft.com/office/2006/relationships/downRev" Target="downrev.xml"/></Relationships>"#).unwrap();
        let blob = zip.finish().unwrap().into_inner();
        let property = record(
            0x13,
            0xf122,
            &[
                0x83a9u16.to_le_bytes().as_slice(),
                &(blob.len() as u32).to_le_bytes(),
                &blob,
            ]
            .concat(),
        );
        let mut info = [0u8; 16];
        info[4..8].copy_from_slice(&7u32.to_le_bytes());
        info[8] = action;
        info[10] = jump;
        let textbox = record(
            15,
            0xf00d,
            &[
                record(0, 3999, &4u32.to_le_bytes()),
                record(0, TEXT_CHARS_ATOM, &[b'a', 0, b'b', 0, b'c', 0]),
                record((instance << 4) | 15, 0x0ff2, &record(0, 0x0ff3, &info)),
                record(
                    range_instance << 4,
                    0x0fdf,
                    &[range.0, 0, 0, 0, range.1, 0, 0, 0],
                ),
                // An origin ruler: left margin and indent 0 at level 0.
                record(
                    0,
                    4006,
                    &[(8u32 | 256).to_le_bytes().as_slice(), &[0; 4]].concat(),
                ),
            ]
            .concat(),
        );
        let local = record(
            15,
            SLIDE_CONTAINER,
            &record(
                15,
                1036,
                &record(
                    15,
                    0xf002,
                    &shape(
                        202,
                        0x200,
                        0xf010,
                        [0, 0, 100, 100],
                        vec![property, textbox],
                    ),
                ),
            ),
        );
        let target: Vec<u8> = "https://example.invalid/"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let list = record(
            15,
            0x0409,
            &record(
                15,
                0x0fd7,
                &[
                    record(0, 0x0fd3, &7u32.to_le_bytes()),
                    record(0x10, 0x0fba, &target),
                ]
                .concat(),
            ),
        );
        let list_offset = local.len();
        let document = [local, list].concat();
        let local_span = record_span_with_end(&document, 0, &mut 20, "slide")
            .unwrap()
            .0;
        let list_span = record_span_with_end(&document, list_offset, &mut 20, "list")
            .unwrap()
            .0;
        let mut p = presentation(local_span);
        p.hyperlinks = media::hyperlink_catalog(&document, &[list_span], &mut 1000);
        slide(
            0,
            &p,
            &document,
            None,
            &mut media::SpanStore::new(Vec::new()),
            &mut 10_000,
            &mut 100_000,
            &mut (16 * 1024 * 1024),
        )
    }

    #[test]
    fn text_interactions_project_click_mouse_over_and_jump_ranges() {
        let links = |slide: Slide| -> Vec<[Option<String>; 3]> {
            let SlideElement::Shape(shape) = &slide.elements[0] else {
                panic!("shape")
            };
            shape.text_body.as_ref().unwrap().paragraphs[0]
                .runs
                .iter()
                .map(|run| match run {
                    TextRun::Text(run) => [
                        run.hyperlink.clone(),
                        run.hyperlink_action.clone(),
                        run.hyperlink_mouse_over.clone(),
                    ],
                    _ => panic!("text run"),
                })
                .collect()
        };
        let url = Some("https://example.invalid/".to_owned());
        // [MS-PPT] 2.6.9: action 4 is a hyperlink, 3 a jump; instance 1 of
        // the container and of its range is the mouse-over interaction.
        assert_eq!(
            links(interactive_text(0, 4, 0, 0, (0, 2)).unwrap()),
            [[url.clone(), None, None], [None, None, None]]
        );
        assert_eq!(
            links(interactive_text(1, 4, 0, 1, (0, 2)).unwrap()),
            [[None, None, url], [None, None, None]]
        );
        assert_eq!(
            links(interactive_text(0, 3, 1, 0, (0, 2)).unwrap()),
            [
                [
                    None,
                    Some("ppaction://hlinkshowjump?jump=nextslide".to_owned()),
                    None
                ],
                [None, None, None]
            ]
        );
        // Actions without a DrawingML run property state nothing.
        for (instance, action, jump) in [(1, 3, 1), (0, 3, 0), (0, 2, 0)] {
            assert_eq!(
                links(interactive_text(instance, action, jump, instance, (0, 2)).unwrap()),
                [[None, None, None]]
            );
        }
        // [MS-PPT] 2.9.59: a range that ends before it begins is not a
        // range, whatever the action does.
        for action in [4, 2] {
            let error = interactive_text(0, action, 0, 0, (2, 1))
                .map(|_| ())
                .unwrap_err();
            assert!(error.contains("ends before it begins"), "{error}");
        }
        // [MS-PPT] 2.9.59: a range past the three characters is not valid
        // for the text, whatever the action does.
        for action in [4, 2] {
            let error = interactive_text(0, action, 0, 0, (0, 4)).unwrap_err();
            assert!(error.contains("exceeds its text"), "{error}");
        }
        assert!(interactive_text(0, 4, 0, 0, (0, 3)).is_ok());
        // A range of the other interaction, or an unknown one, is malformed.
        for (instance, range) in [(0, 1), (1, 0), (2, 2)] {
            assert!(interactive_text(instance, 4, 0, range, (0, 2)).is_err());
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
    fn ole_objects_draw_their_stored_picture_or_a_placeholder_frame() {
        use crate::rejection::{take_best_effort, BestEffort};
        let list = |objects: &[(u16, u32, u32, u32)]| Some(ex_obj_list(objects));
        let picture = || vec![properties(&[(0x4104, 1)]), ole_ref(7)];
        // A linked object shows its stored picture, as an embedded one does.
        take_best_effort();
        let model = project(75, OLE, picture(), png_blip(), list(&[(0x0fce, 1, 1, 7)])).unwrap();
        assert!(matches!(
            &model.elements[..],
            [SlideElement::Picture(picture)] if picture.image_path == "legacy-ppt/image/1"
        ));
        assert_eq!(take_best_effort(), []);
        // Aspects and controls without Office evidence draw it too.
        for (objects, site) in [
            (
                list(&[(0x0fcc, 4, 0, 7)]),
                "PowerPoint OLE object drawn as icon or thumbnail",
            ),
            (list(&[(0x0fee, 1, 2, 7)]), "PowerPoint ActiveX control"),
        ] {
            let model = project(75, OLE, picture(), png_blip(), objects).unwrap();
            assert!(matches!(
                &model.elements[..],
                [SlideElement::Picture(picture)] if picture.image_path == "legacy-ppt/image/1"
            ));
            assert_eq!(take_best_effort(), [(BestEffort::Approximated, site)]);
        }
        // No picture frame BLIP to show: an unfilled frame keeps the place.
        for (kind, children) in [
            (75, vec![ole_ref(7)]),
            (1, vec![properties(&[(0x4104, 1)]), ole_ref(7)]),
        ] {
            let model =
                project(kind, OLE, children, png_blip(), list(&[(0x0fcc, 1, 0, 7)])).unwrap();
            let [SlideElement::Shape(frame)] = &model.elements[..] else {
                panic!("placeholder frame: {:?}", model.elements)
            };
            assert_eq!(frame.geometry, "rect");
            assert_eq!((frame.width, frame.height), (158750, 158750));
            assert!(matches!(frame.fill, Some(Fill::None)));
            assert_eq!(frame.stroke.as_ref().unwrap().color, "808080");
            assert_eq!(
                take_best_effort(),
                [(
                    BestEffort::Approximated,
                    "PowerPoint OLE object has no presentation picture"
                )]
            );
        }
        for (kind, children, objects, expected) in [
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
    fn pict_picture_is_retained_as_an_explicit_unsupported_image() {
        // A minimal v2 picture in an uncompressed MS-ODRAW 2.2.26 BLIP.
        let pict = [0, 0, 0, 0, 0, 0, 0, 20, 0, 30, 0, 17, 2, 255, 0, 255];
        let mut payload = vec![0; 16 + 34];
        payload[16..20].copy_from_slice(&(pict.len() as u32).to_le_bytes());
        payload[44..48].copy_from_slice(&(pict.len() as u32).to_le_bytes());
        payload[48..50].fill(0xfe);
        payload.extend(pict);
        let blip = record(0x5420, 0xf01c, &payload);
        let model = project(
            75,
            0x200,
            vec![properties(&[(0x4104, 1)])],
            blip.clone(),
            None,
        )
        .unwrap();
        let SlideElement::Picture(picture) = &model.elements[0] else {
            panic!("picture")
        };
        assert_eq!(picture.mime_type, "image/x-pict");
        assert_eq!(picture.image_path, "legacy-ppt/image/1");
        assert!(picture.width > 0 && picture.height > 0);
        // Recovery is picture-frame policy, not implicit acceptance of fills.
        let fill = properties(&[(0x180, 3), (0x4186, 1), (0x1bf, 0x0010_0010)]);
        assert!(project(1, 0x200, vec![fill.clone()], blip.clone(), None)
            .unwrap_err()
            .contains("fill BLIP"));
        // A JPEG frame without a decoder (ITU-T T.81 SOF9, arithmetic
        // coding) keeps the picture's place as a frame with its own line.
        use crate::rejection::{take_best_effort, BestEffort};
        let jpeg = [0xff, 0xd8, 0xff, 0xc9, 0, 11, 8, 0, 2, 0, 3, 1, 1, 0x11, 0];
        let jpeg = record(0x46a0, 0xf01d, &[vec![0; 17], jpeg.to_vec()].concat());
        take_best_effort();
        let lined = properties(&[(0x4104, 1), (0x1c0, 0x0000ff), (0x1ff, 0x0008_0008)]);
        let model = project(75, 0x200, vec![lined], jpeg.clone(), None).unwrap();
        let [SlideElement::Shape(frame)] = &model.elements[..] else {
            panic!("placeholder frame: {:?}", model.elements)
        };
        assert_eq!((frame.geometry.as_str(), frame.width), ("rect", 158750));
        assert!(matches!(frame.fill, Some(Fill::None)));
        assert_eq!(frame.stroke.as_ref().unwrap().color, "FF0000");
        assert_eq!(
            take_best_effort(),
            [(
                BestEffort::Approximated,
                "unsupported OfficeArt JPEG frame encoding"
            )]
        );
        assert!(project(1, 0x200, vec![fill], jpeg, None)
            .unwrap_err()
            .contains("fill BLIP"));
        let mut broken = blip;
        broken[8 + 44..8 + 48].copy_from_slice(&999u32.to_le_bytes());
        assert!(project(75, 0x200, vec![properties(&[(0x4104, 1)])], broken, None).is_err());
    }

    #[test]
    fn picture_adjustments_and_blip_fills_project_or_reject() {
        use crate::rejection::{take_best_effort, BestEffort};
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
            (vec![(0x109, 0x8000), (0x13f, 0x0002_0002)], "brightness"),
            (vec![(0x108, 0x4ccd), (0x107, 0x00ff_ffff)], "brightness"),
            (
                vec![(0x107, 0x0800_0001)],
                "indexed picture transparent color",
            ),
            (vec![(0x11a, 0x0000ff)], "recolor"),
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
        // fPictureGray alone is a gray picture; with brightness the gray
        // conversion comes after the luminance change.
        for (values, expected) in [
            (vec![(0x13f, 0x0004_0004)], vec![BlipEffect::Grayscale]),
            (
                vec![(0x109, 0x4000), (0x13f, 0x0004_0004)],
                vec![
                    BlipEffect::Luminance {
                        bright: 0.5,
                        contrast: 0.0,
                    },
                    BlipEffect::Grayscale,
                ],
            ),
        ] {
            let values = [vec![(0x4104, 1)], values].concat();
            let model = project(75, 0x200, vec![properties(&values)], png_blip(), None).unwrap();
            let SlideElement::Picture(picture) = &model.elements[0] else {
                panic!("picture")
            };
            assert_eq!(picture.blip_effects, expected);
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
        // Non-plain pattern and picture fills are not left unfilled.
        for (values, expected) in [
            (
                vec![(0x180, 1), (0x4186, 1), (0x1bf, 0x0002_0002)],
                Some("pattern fill is not projected"),
            ),
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
        // A frame with another fill kind projects it as a shape would.
        take_best_effort();
        assert!(matches!(
            picture_fill(backing(&[(0x180, 4), (0x181, 0xff), (0x1bf, 0x0010_0010)]).unwrap()),
            Some(Fill::Gradient { ref stops, .. })
                if stops.iter().any(|stop| stop.color == "FF0000")
        ));
        assert_eq!(
            take_best_effort(),
            [(
                BestEffort::Approximated,
                "PowerPoint picture frame non-solid fill"
            )]
        );
        // A pattern or texture behind a picture is a rectangle under it:
        // the renderer paints no picture fill behind a picture.
        for kind in [1, 2] {
            let model = backing(&[(0x180, kind), (0x4186, 1), (0x1bf, 0x0010_0010)]).unwrap();
            let [SlideElement::Shape(under), SlideElement::Picture(picture)] = &model.elements[..]
            else {
                panic!("backing rectangle and picture: {:?}", model.elements)
            };
            assert!(picture.fill.is_none());
            assert!(matches!(
                under.fill,
                Some(Fill::Image { tile: Some(_), .. })
            ));
            assert_eq!(
                (
                    under.geometry.as_str(),
                    under.id.as_deref(),
                    under.stroke.is_none()
                ),
                ("rect", None, true)
            );
            assert_eq!(
                (under.x, under.y, under.width, under.height),
                (picture.x, picture.y, picture.width, picture.height)
            );
            take_best_effort();
        }
        // msofillTexture tiles the BLIP at its own size from the top left.
        let texture = [(0x180, 2), (0x4186, 1), (0x1bf, 0x0060_0060)];
        let model = project(1, 0x200, vec![properties(&texture)], png_blip(), None).unwrap();
        let SlideElement::Shape(shape) = &model.elements[0] else {
            panic!("shape")
        };
        let Some(Fill::Image {
            stretch: false,
            dpi: Some(0),
            tile: Some(tile),
            rot_with_shape: Some(true),
            duotone: None,
            ..
        }) = &shape.fill
        else {
            panic!("texture fill: {:?}", shape.fill)
        };
        assert_eq!(
            (tile.tx, tile.ty, tile.sx, tile.sy, tile.algn.as_deref()),
            (Some(0), Some(0), Some(1.0), Some(1.0), Some("tl"))
        );
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
        // An 8x8 bitmap tiles whole at 96 DPI; each opacity applies to its
        // own colour. Another size has no evidence and tiles like it.
        let translucent = [pattern.to_vec(), vec![(0x182, 0x8000), (0x184, 0x4000)]].concat();
        for (values, blip, flags, colors, note) in [
            (
                pattern.to_vec(),
                gif_blip(8, 8),
                0x200,
                ("0B0B0D", "217EA8"),
                false,
            ),
            (
                translucent,
                gif_blip(8, 8),
                0x200,
                ("0B0B0D40", "217EA880"),
                false,
            ),
            (
                pattern.to_vec(),
                png_blip(),
                0x200,
                ("0B0B0D", "217EA8"),
                true,
            ),
        ] {
            take_best_effort();
            let model = project(1, flags, vec![properties(&values)], blip, None).unwrap();
            let SlideElement::Shape(shape) = &model.elements[0] else {
                panic!("shape")
            };
            let Some(Fill::Image {
                dpi: Some(96),
                src_rect: None,
                rot_with_shape: Some(false),
                tile: Some(_),
                duotone: Some(duotone),
                ..
            }) = &shape.fill
            else {
                panic!("pattern fill: {:?}", shape.fill)
            };
            assert_eq!((duotone.clr1.as_str(), duotone.clr2.as_str()), colors);
            assert_eq!(
                take_best_effort(),
                if note {
                    vec![(
                        BestEffort::Approximated,
                        "PowerPoint pattern bitmap size is not projected",
                    )]
                } else {
                    Vec::new()
                }
            );
        }
        // A flipped shape tiles the bitmap as seen through its reflection,
        // offset so that a tile starts at the edge drawn on the left: the
        // renderer reflects the fill with the shape, PowerPoint does not.
        let cells = [0xcc, 0x66, 0x33, 0x99, 0xcc, 0x66, 0x33, 0x99];
        let stored = record(
            0x6e00,
            0xf01e,
            &[vec![0; 17], crate::ppt::pattern::png(cells)].concat(),
        );
        for (flags, offset) in [
            (0x240, (6350, 0)),
            (0x280, (0, 6350)),
            (0x2c0, (6350, 6350)),
        ] {
            take_best_effort();
            let model =
                project(1, flags, vec![properties(&pattern)], stored.clone(), None).unwrap();
            let SlideElement::Shape(shape) = &model.elements[0] else {
                panic!("shape")
            };
            let Some(Fill::Image {
                image_path,
                mime_type,
                dpi: Some(96),
                src_rect: None,
                tile: Some(tile),
                duotone: Some(_),
                ..
            }) = &shape.fill
            else {
                panic!("pattern fill: {:?}", shape.fill)
            };
            // The written bitmap is the session's first image after the
            // stored one. The shape is 158750 EMU wide and high; a tile is
            // 8 pixels at 96 DPI, 76200 EMU.
            assert_eq!(
                (image_path.as_str(), mime_type.as_str()),
                ("legacy-ppt/image/2", "image/png")
            );
            assert_eq!((tile.tx, tile.ty), (Some(offset.0), Some(offset.1)));
            assert_eq!(take_best_effort(), []);
        }
        // A bitmap this reader cannot rewrite tiles reflected.
        let model = project(1, 0x240, vec![properties(&pattern)], gif_blip(8, 8), None).unwrap();
        assert!(matches!(
            &model.elements[0],
            SlideElement::Shape(shape) if matches!(
                &shape.fill,
                Some(Fill::Image { image_path, .. }) if image_path == "legacy-ppt/image/1"
            )
        ));
        assert_eq!(
            take_best_effort(),
            [(
                BestEffort::Approximated,
                "PowerPoint pattern fill on a flipped shape is not projected"
            )]
        );
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
