//! Direct, bounded projection of the validated binary slide tree into the
//! presentation renderer model. No package or XML intermediary is involved.
use super::*;
use ooxml_common::blip::SrcRect;
use pptx_model::{
    Fill, PictureElement, ShapeElement, Slide, SlideElement, SlideElementOrigin,
    SlideElementSource, TextBody,
};

const MAX_MODEL_SHAPES: usize = 100_000;
// Bounded upper charge for model-owned paint, identifiers and media path/MIME
// strings. Text and custom-path backing are charged by their own projectors.
const MODEL_SHAPE_STRINGS_BYTES: usize = 384;

fn document_text_axes(
    text_type: Option<u16>,
    placeholder: bool,
    master_linked: bool,
    outline: bool,
    axes: Option<text_style::ParagraphAxes>,
) -> Option<text_style::ParagraphAxes> {
    (text_type == Some(4) && !placeholder && !master_linked && !outline)
        .then_some(axes)
        .flatten()
}

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
                    .map(|value| value.to_model(self.model_budget))
                    .transpose()?;
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
            if group.omitted() || group.props.hidden || (inherited && group.is_placeholder()) {
                return Ok(());
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
            if shape.omitted() || shape.props.hidden || (inherited && shape.is_placeholder()) {
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
        let transform = direct_transform::flatten(direct_transform::leaf(&shape)?, ancestors);
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
        if shape.kind == 75 && shape.props.picture != 0 {
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
                let (_, stroke) = paint.model_with_custom_geometry(
                    self.presentation.schemes[self.index].as_ref(),
                    false,
                    true,
                    None,
                );
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
        }
        let text = self.text_body(&shape, inherited)?;
        let preset = paint.geometry(shape.kind);
        if text.is_none() && preset.is_none() && custom.is_none() {
            return Ok(());
        }
        self.charge_shape_strings()?;
        let (geometry_name, paths, allow_fill, allow_line) = match custom {
            Some(g) => ("custGeom".to_owned(), Some(g.paths), g.fill, g.stroke),
            None => (
                preset.unwrap_or("rect").to_owned(),
                None,
                preset.is_some() && !matches!(shape.kind, 20 | 32),
                true,
            ),
        };
        let image = if allow_fill {
            paint
                .foreground_image()
                .map(|(id, alpha, rotate)| self.image_fill(id, alpha, rotate))
                .transpose()?
                .flatten()
        } else {
            None
        };
        let (fill, stroke) = paint.model_with_custom_geometry(
            self.presentation.schemes[self.index].as_ref(),
            allow_fill,
            allow_line,
            image,
        );
        let gradient_fill = if shape.props.rotation == 0
            && ancestors
                .iter()
                .all(|group| group.rot == 0.0 && !group.flip_h && !group.flip_v)
        {
            paint
                .project_gradient(
                    &gradient,
                    allow_fill,
                    self.presentation.schemes[self.index].as_ref(),
                    self.work_budget,
                    self.model_budget,
                )?
                .map(|value| value.to_model(self.model_budget))
                .transpose()?
        } else {
            None
        };
        let fill = gradient_fill.or(fill);
        self.push(
            SlideElement::Shape(ShapeElement {
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
                id: Some(shape.id.to_string()),
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

    fn image_fill(&mut self, id: u32, opacity: u32, rotate: bool) -> Result<Option<Fill>, String> {
        if !self
            .media
            .reference(id, self.backing, self.pictures, self.work_budget)?
        {
            return Ok(None);
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
        }))
    }

    fn text_body(
        &mut self,
        shape: &SpannedShape,
        inherited: bool,
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
        let levels = linked.or_else(|| {
            shape
                .is_placeholder()
                .then(|| {
                    self.presentation.text_masters[self.index]
                        .as_deref()
                        .and_then(|m| text_type.and_then(|t| m.levels(t)))
                })
                .flatten()
        });
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
                auto_number: None,
            },
            text_style::direct_model::DirectAxes {
                ruler: local_ruler,
                document: document_text_axes(
                    text_type,
                    shape.is_placeholder(),
                    master_link.is_some(),
                    outline_body,
                    self.presentation.document_text_axes,
                ),
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
            vert: if p.text_flow == Some(1) && p.font_direction.unwrap_or(0) == 0 {
                "eaVert"
            } else {
                "horz"
            }
            .to_owned(),
            auto_fit: "none".to_owned(),
            font_scale: None,
            ln_spc_reduction: None,
            num_col: 1,
            spc_col: 0,
            rtl_col: false,
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
            document_text_axes: None,
            fonts: Vec::new(),
            schemes: vec![None],
            image_entries: Vec::new(),
            backgrounds: vec![None],
            object_masters: vec![std::rc::Rc::from([])],
            size: (720, 540),
        }
    }

    #[test]
    fn document_type4_origins_are_limited_to_unlinked_freeform_text() {
        let axes = Some(text_style::ParagraphAxes {
            margin: Some(180),
            indent: Some(90),
        });
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
            &mut 100_000,
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
}
