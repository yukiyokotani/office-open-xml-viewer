//! Direct-model projection of already validated floating pictures and
//! drawing shapes of the main and header documents.

use super::{shape, unsupported, Content, Mode, Part, ResolvedDrawing, Store};
use crate::doc::pictures::DirectPictureResource;
use docx_model::{
    AnchorAcquisitionWire, AnchorAxisChoiceWire, AnchorAxisWire, AnchorBehaviorWire,
    AnchorEdgesWire, AnchorExtentWire, AnchorGroupWire, AnchorResolvedChildFrameWire,
    AnchorSimplePositionWire, AnchorValueStatusWire, AnchorWrapKindWire, AnchorWrapWire, ImageRun,
    LineEnd, PathCmd, ShapeFill, ShapeRun,
};

#[cfg(test)]
#[derive(Debug)]
pub(in crate::doc) struct DirectFloatingPicture {
    pub image: ImageRun,
    pub occurrence_id: String,
}

#[derive(Debug)]
pub(in crate::doc) struct DirectFloatingShape {
    pub shape: ShapeRun,
    /// One-based FTXBXS index whose text fills `shape.text_box_content`.
    pub text: Option<usize>,
    /// The shape's spid, which the FTXBXS `lid` must name.
    pub spid: u32,
}

/// One drawing occurrence: a single picture or shape, or the members of a
/// group, all sharing one anchor host (as DOCX `wpg:wgp` members do).
#[derive(Debug)]
pub(in crate::doc) struct DirectFloating {
    pub occurrence_id: String,
    pub runs: Vec<DirectRun>,
}

#[derive(Debug)]
pub(in crate::doc) enum DirectRun {
    Image(Box<ImageRun>),
    Shape(Box<DirectFloatingShape>),
}

/// Where one run sits inside its drawing frame: the whole frame for a single
/// drawing, or a mapped group member frame (x, y, width, height in EMUs).
struct Placed {
    frame: [f64; 4],
    rotation: f64,
    flip: [bool; 2],
    group: Option<(usize, usize)>,
}

fn wrap_mode(wrapping: u8) -> &'static str {
    match wrapping {
        1 => "topAndBottom",
        2 => "square",
        3 => "none",
        _ => unreachable!("unsupported wrap modes are rejected during resolution"),
    }
}

impl Store<'_> {
    pub(in crate::doc) fn has_selected_direct_resources(&self) -> bool {
        !self.selected_images.is_empty()
    }

    /// Project the drawing anchored at `cp` of `part`'s story.
    pub(in crate::doc) fn direct_drawing(
        &mut self,
        part: Part,
        cp: usize,
        remaining_bytes: &mut usize,
    ) -> Result<Option<DirectFloating>, String> {
        let Some(facts) = self.resolve(part, cp, Mode::Direct)? else {
            return Ok(None);
        };
        let occurrence_id = format!("legacy-doc-float-{}", facts.occurrence);
        *remaining_bytes = remaining_bytes
            .checked_sub(occurrence_id.capacity())
            .ok_or("OUTPUT_TOO_LARGE")?;
        let whole = Placed {
            frame: [0.0, 0.0, facts.extent[0] as f64, facts.extent[1] as f64],
            rotation: 0.0,
            flip: facts.flip,
            group: None,
        };
        let mut runs = Vec::new();
        let push = |runs: &mut Vec<DirectRun>, run: DirectRun, remaining: &mut usize| {
            *remaining = remaining
                .checked_sub(std::mem::size_of::<DirectRun>())
                .ok_or("OUTPUT_TOO_LARGE")?;
            runs.push(run);
            Ok::<(), String>(())
        };
        match &facts.content {
            Content::Picture {
                image_index,
                extension,
                crop,
            } => {
                let image = self.direct_image(
                    &facts,
                    (*image_index, extension, *crop),
                    &whole,
                    &occurrence_id,
                    remaining_bytes,
                )?;
                push(
                    &mut runs,
                    DirectRun::Image(Box::new(image)),
                    remaining_bytes,
                )?;
            }
            Content::Shape(shape) => {
                let fill = self.direct_fill(shape)?;
                let shape = direct_shape(
                    fill,
                    &facts,
                    shape,
                    &whole,
                    facts.shape_id,
                    &occurrence_id,
                    remaining_bytes,
                )?;
                push(
                    &mut runs,
                    DirectRun::Shape(Box::new(shape)),
                    remaining_bytes,
                )?;
            }
            Content::Group(members) => {
                for (index, member) in members.iter().enumerate() {
                    let placed = Placed {
                        frame: member.frame,
                        rotation: member.rotation,
                        flip: member.flip,
                        group: Some((index, members.len())),
                    };
                    let run = match &member.content {
                        Content::Picture {
                            image_index,
                            extension,
                            crop,
                        } => DirectRun::Image(Box::new(self.direct_image(
                            &facts,
                            (*image_index, extension, *crop),
                            &placed,
                            &occurrence_id,
                            remaining_bytes,
                        )?)),
                        Content::Shape(shape) => DirectRun::Shape(Box::new(direct_shape(
                            self.direct_fill(shape)?,
                            &facts,
                            shape,
                            &placed,
                            member.spid,
                            &occurrence_id,
                            remaining_bytes,
                        )?)),
                        Content::Group(_) => unreachable!("groups are flattened"),
                    };
                    push(&mut runs, run, remaining_bytes)?;
                }
            }
        }
        Ok(Some(DirectFloating {
            occurrence_id,
            runs,
        }))
    }

    #[cfg(test)]
    pub(in crate::doc) fn direct_picture(
        &mut self,
        cp: usize,
        remaining_bytes: &mut usize,
    ) -> Result<Option<DirectFloatingPicture>, String> {
        let Some(mut drawing) = self.direct_drawing(Part::Main, cp, remaining_bytes)? else {
            return Ok(None);
        };
        assert_eq!(drawing.runs.len(), 1);
        let DirectRun::Image(image) = drawing.runs.pop().unwrap() else {
            panic!("expected a picture");
        };
        Ok(Some(DirectFloatingPicture {
            image: *image,
            occurrence_id: drawing.occurrence_id,
        }))
    }

    /// The shape's paint as a DOCX fill: solid, or the stretched picture of an
    /// msofillPicture fill, sharing the floating picture resources (ECMA-376
    /// 20.1.8.14 blipFill with a:stretch/a:fillRect).
    fn direct_fill(&mut self, shape: &shape::Facts) -> Result<Option<ShapeFill>, String> {
        if let Some((index, _)) = shape.fill_picture {
            let extension = self.images[&index]
                .as_ref()
                .expect("loaded picture fill")
                .extension;
            self.selected_images.insert(index);
            return Ok(Some(ShapeFill::Image {
                image_path: format!("legacy-doc/float/{index}"),
                mime_type: mime(extension)?.to_string(),
                svg_image_path: None,
                src_rect: None,
                fill_rect: None,
                tile: None,
                alpha: None,
                duotone: None,
            }));
        }
        Ok(shape.fill.clone().map(|color| ShapeFill::Solid { color }))
    }

    fn direct_image(
        &mut self,
        facts: &ResolvedDrawing,
        (image_index, extension, crop): (usize, &str, [i64; 4]),
        placed: &Placed,
        occurrence_id: &str,
        remaining_bytes: &mut usize,
    ) -> Result<ImageRun, String> {
        let mime_type = mime(extension)?.to_string();
        let image_path = format!("legacy-doc/float/{image_index}");
        let acquisition = acquisition(facts, occurrence_id.into(), placed);
        let [top, bottom, left, right] = crop;
        let image = ImageRun {
            image_path,
            mime_type,
            svg_image_path: None,
            src_rect: (crop != [0; 4]).then_some(ooxml_common::blip::SrcRect {
                l: left as f64 / 100_000.0,
                t: top as f64 / 100_000.0,
                r: right as f64 / 100_000.0,
                b: bottom as f64 / 100_000.0,
            }),
            width_pt: placed.frame[2] / 12_700.0,
            height_pt: placed.frame[3] / 12_700.0,
            rotation: 0.0,
            flip_h: placed.flip[0],
            flip_v: placed.flip[1],
            anchor: true,
            anchor_x_pt: (facts.x_emu as f64 + placed.frame[0]) / 12_700.0,
            anchor_y_pt: (facts.y_emu as f64 + placed.frame[1]) / 12_700.0,
            anchor_x_from_margin: matches!(facts.horizontal, "margin" | "column"),
            anchor_y_from_para: facts.vertical == "paragraph",
            color_replace_from: None,
            duotone: None,
            alpha: None,
            wrap_mode: Some(wrap_mode(facts.wrapping).into()),
            dist_top: facts.distances[1] as f64 / 12_700.0,
            dist_bottom: facts.distances[3] as f64 / 12_700.0,
            dist_left: facts.distances[0] as f64 / 12_700.0,
            dist_right: facts.distances[2] as f64 / 12_700.0,
            wrap_side: (facts.wrapping == 2).then(|| facts.side.into()),
            allow_overlap: facts.overlap,
            anchor_x_align: facts.align[0].map(str::to_owned),
            anchor_y_align: facts.align[1].map(str::to_owned),
            anchor_x_relative_from: Some(facts.horizontal.into()),
            anchor_y_relative_from: Some(facts.vertical.into()),
            anchor_acquisition: Some(acquisition),
        };
        let required = image_payload(&image)?;
        *remaining_bytes = remaining_bytes
            .checked_sub(required)
            .ok_or("OUTPUT_TOO_LARGE")?;
        self.selected_images.insert(image_index);
        Ok(image)
    }

    pub(in crate::doc) fn append_referenced_direct_resources(
        self,
        resources: &mut Vec<DirectPictureResource>,
        references: &[&str],
        remaining_bytes: &mut usize,
    ) -> Result<(), String> {
        for reference in references {
            let Some(suffix) = reference.strip_prefix("legacy-doc/float/") else {
                continue;
            };
            let index = suffix
                .parse::<usize>()
                .map_err(|_| unsupported("invalid direct DOC floating picture resource key"))?;
            if format!("legacy-doc/float/{index}") != *reference
                || !self.selected_images.contains(&index)
                || self.images.get(&index).and_then(Option::as_ref).is_none()
            {
                return Err(unsupported("dangling direct DOC floating picture resource"));
            }
        }
        self.append_direct_resources_with(
            resources,
            |index| {
                let value = format!("legacy-doc/float/{index}");
                references.binary_search(&value.as_str()).is_ok()
            },
            remaining_bytes,
        )
    }

    #[cfg(test)]
    pub(in crate::doc) fn append_direct_resources(
        self,
        resources: &mut Vec<DirectPictureResource>,
        remaining_bytes: &mut usize,
    ) -> Result<(), String> {
        let selected = self.selected_images.clone();
        self.append_direct_resources_with(
            resources,
            |index| selected.contains(&index),
            remaining_bytes,
        )
    }

    fn append_direct_resources_with(
        self,
        resources: &mut Vec<DirectPictureResource>,
        selected: impl Fn(usize) -> bool,
        remaining_bytes: &mut usize,
    ) -> Result<(), String> {
        let count = self
            .selected_images
            .iter()
            .filter(|index| selected(**index))
            .count();
        let old_capacity = resources.capacity();
        let minimum_capacity = resources
            .len()
            .checked_add(count)
            .ok_or("OUTPUT_TOO_LARGE")?;
        let minimum_added = minimum_capacity.saturating_sub(old_capacity);
        let mut required = minimum_added
            .checked_mul(std::mem::size_of::<DirectPictureResource>())
            .ok_or("OUTPUT_TOO_LARGE")?;
        for index in self
            .selected_images
            .iter()
            .filter(|index| selected(**index))
        {
            let image = self.images[index]
                .as_ref()
                .expect("selected floating image");
            required = required
                .checked_add(format!("legacy-doc/float/{index}").capacity())
                .and_then(|v| {
                    v.checked_add(match &image.bytes {
                        std::borrow::Cow::Borrowed(bytes) => bytes.len(),
                        std::borrow::Cow::Owned(bytes) => bytes.capacity(),
                    })
                })
                .ok_or("OUTPUT_TOO_LARGE")?;
        }
        *remaining_bytes = remaining_bytes
            .checked_sub(required)
            .ok_or("OUTPUT_TOO_LARGE")?;
        resources
            .try_reserve_exact(count)
            .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
        let added = resources.capacity().saturating_sub(old_capacity);
        *remaining_bytes = remaining_bytes
            .checked_sub(
                added
                    .saturating_sub(minimum_added)
                    .checked_mul(std::mem::size_of::<DirectPictureResource>())
                    .ok_or("OUTPUT_TOO_LARGE")?,
            )
            .ok_or("OUTPUT_TOO_LARGE")?;
        let Store {
            images,
            selected_images,
            ..
        } = self;
        for (index, image) in images {
            if !selected_images.contains(&index) || !selected(index) {
                continue;
            }
            let image = image.expect("selected floating image");
            let key = format!("legacy-doc/float/{index}");
            let bytes = match image.bytes {
                std::borrow::Cow::Owned(bytes) => bytes,
                std::borrow::Cow::Borrowed(source) => {
                    let mut bytes = Vec::new();
                    bytes
                        .try_reserve_exact(source.len())
                        .map_err(|_| "OUTPUT_TOO_LARGE")?;
                    *remaining_bytes = remaining_bytes
                        .checked_sub(bytes.capacity().saturating_sub(source.len()))
                        .ok_or("OUTPUT_TOO_LARGE")?;
                    bytes.extend_from_slice(source);
                    bytes
                }
            };
            resources.push(DirectPictureResource {
                key,
                mime_type: mime(image.extension)?,
                bytes,
            });
        }
        Ok(())
    }
}

fn valid_edges(values: [u32; 4]) -> AnchorEdgesWire {
    let [left, top, right, bottom] = values.map(|v| v as f64 / 12_700.0);
    AnchorEdgesWire {
        top_pt: Some(top),
        top_status: AnchorValueStatusWire::Valid,
        right_pt: Some(right),
        right_status: AnchorValueStatusWire::Valid,
        bottom_pt: Some(bottom),
        bottom_status: AnchorValueStatusWire::Valid,
        left_pt: Some(left),
        left_status: AnchorValueStatusWire::Valid,
    }
}

fn acquisition(
    f: &ResolvedDrawing,
    occurrence_id: String,
    placed: &Placed,
) -> AnchorAcquisitionWire {
    let axis = |relative: &'static str, value: i64, align: Option<&'static str>| AnchorAxisWire {
        relative_from: Some(relative.into()),
        relative_from_status: AnchorValueStatusWire::Valid,
        choice: match align {
            Some(value) => AnchorAxisChoiceWire::Align {
                value: value.into(),
            },
            None => AnchorAxisChoiceWire::Offset {
                value_pt: value as f64 / 12_700.0,
            },
        },
    };
    AnchorAcquisitionWire {
        simple_position: AnchorSimplePositionWire {
            enabled: Some(false),
            status: AnchorValueStatusWire::Valid,
            x_pt: Some(0.0),
            x_status: AnchorValueStatusWire::Valid,
            y_pt: Some(0.0),
            y_status: AnchorValueStatusWire::Valid,
        },
        horizontal: axis(f.horizontal, f.x_emu, f.align[0]),
        vertical: axis(f.vertical, f.y_emu, f.align[1]),
        extent: AnchorExtentWire {
            width_pt: Some(f.extent[0] as f64 / 12_700.0),
            height_pt: Some(f.extent[1] as f64 / 12_700.0),
            width_status: AnchorValueStatusWire::Valid,
            height_status: AnchorValueStatusWire::Valid,
        },
        anchor_distances: valid_edges(f.distances),
        wrap: AnchorWrapWire {
            kind: match f.wrapping {
                1 => AnchorWrapKindWire::TopAndBottom,
                2 => AnchorWrapKindWire::Square,
                _ => AnchorWrapKindWire::None,
            },
            authored_kinds: vec![match f.wrapping {
                1 => "wrapTopAndBottom",
                2 => "wrapSquare",
                _ => "wrapNone",
            }
            .into()],
            side: (f.wrapping == 2).then(|| f.side.into()),
            ..AnchorWrapWire::default()
        },
        behavior: AnchorBehaviorWire {
            behind_doc: Some(f.behind),
            behind_doc_status: AnchorValueStatusWire::Valid,
            relative_height: Some(f.z_order),
            relative_height_status: AnchorValueStatusWire::Valid,
            locked: Some(f.locked),
            locked_status: AnchorValueStatusWire::Valid,
            allow_overlap: Some(f.overlap),
            allow_overlap_status: AnchorValueStatusWire::Valid,
            layout_in_cell: Some(f.in_cell),
            layout_in_cell_status: AnchorValueStatusWire::Valid,
        },
        // Group members share the group's anchor frame and extent; their
        // own frame is resolved relative to it (MS-ODRAW 2.2.38-2.2.39).
        group: placed.group.map(|(index, count)| AnchorGroupWire {
            child_source_id: format!("{occurrence_id}-member-{index}"),
            source_index: index,
            source_count: count,
            transform_chain: Vec::new(),
            child_transform: None,
            resolved_child_frame: AnchorResolvedChildFrameWire {
                offset_x_pt: placed.frame[0] / 12_700.0,
                offset_y_pt: placed.frame[1] / 12_700.0,
                width_pt: placed.frame[2] / 12_700.0,
                height_pt: placed.frame[3] / 12_700.0,
                rotation_deg: placed.rotation,
                flip_h: placed.flip[0],
                flip_v: placed.flip[1],
            },
        }),
        occurrence_id,
        ..AnchorAcquisitionWire::default()
    }
}

/// Checked sum of retained heap payload plus a fixed struct size.
struct Payload(usize);

impl Payload {
    fn add(&mut self, bytes: usize) -> Result<(), String> {
        self.0 = self.0.checked_add(bytes).ok_or("OUTPUT_TOO_LARGE")?;
        Ok(())
    }
    fn strings<'s>(
        &mut self,
        values: impl IntoIterator<Item = Option<&'s String>>,
    ) -> Result<(), String> {
        for value in values.into_iter().flatten() {
            self.add(value.capacity())?;
        }
        Ok(())
    }
    fn acquisition(&mut self, facts: &AnchorAcquisitionWire) -> Result<(), String> {
        self.add(facts.occurrence_id.capacity())?;
        fn align(axis: &AnchorAxisWire) -> Option<&String> {
            match &axis.choice {
                AnchorAxisChoiceWire::Align { value } => Some(value),
                _ => None,
            }
        }
        self.strings([
            facts.horizontal.relative_from.as_ref(),
            facts.vertical.relative_from.as_ref(),
            align(&facts.horizontal),
            align(&facts.vertical),
            facts.wrap.side.as_ref(),
        ])?;
        self.add(
            facts
                .wrap
                .authored_kinds
                .capacity()
                .checked_mul(std::mem::size_of::<String>())
                .ok_or("OUTPUT_TOO_LARGE")?,
        )?;
        self.strings(facts.wrap.authored_kinds.iter().map(Some))?;
        if let Some(group) = &facts.group {
            self.add(group.child_source_id.capacity())?;
        }
        for axis in [
            &facts.relative_size.horizontal,
            &facts.relative_size.vertical,
        ]
        .into_iter()
        .flatten()
        {
            self.strings([axis.relative_from.as_ref()])?;
        }
        Ok(())
    }
}

fn image_payload(image: &ImageRun) -> Result<usize, String> {
    let mut total = Payload(std::mem::size_of::<ImageRun>());
    total.strings([
        Some(&image.image_path),
        Some(&image.mime_type),
        image.wrap_mode.as_ref(),
        image.wrap_side.as_ref(),
        image.anchor_x_align.as_ref(),
        image.anchor_y_align.as_ref(),
        image.anchor_x_relative_from.as_ref(),
        image.anchor_y_relative_from.as_ref(),
    ])?;
    total.acquisition(
        image
            .anchor_acquisition
            .as_ref()
            .expect("floating acquisition"),
    )?;
    Ok(total.0)
}

/// Project resolved drawing-shape facts onto the DOCX shape model: the same
/// anchor/wrap facts as pictures, an ECMA-376 preset outline, solid paint and
/// text box margins (ECMA-376 20.1.9.18, 20.1.8.54, 20.1.2.2.24 and
/// 20.4.2.3; bodyPr lIns/tIns/rIns/bIns and spAutoFit, 21.1.2.1.1-2).
fn direct_shape(
    fill: Option<ShapeFill>,
    facts: &ResolvedDrawing,
    shape: &shape::Facts,
    placed: &Placed,
    spid: u32,
    occurrence_id: &str,
    remaining_bytes: &mut usize,
) -> Result<DirectFloatingShape, String> {
    let mut acquisition = acquisition(facts, occurrence_id.into(), placed);
    let member = placed.group.map_or(0, |(index, _)| index as u32);
    let [relative_width, relative_height] = shape.relative_size;
    let axis = |size: Option<(f64, &'static str)>| {
        size.map(|(fraction, from)| docx_model::AnchorRelativeSizeAxisWire {
            relative_from: Some(from.into()),
            relative_from_status: AnchorValueStatusWire::Valid,
            fraction: Some(fraction),
            fraction_status: AnchorValueStatusWire::Valid,
        })
    };
    acquisition.relative_size = docx_model::AnchorRelativeSizeWire {
        horizontal: axis(relative_width),
        vertical: axis(relative_height),
    };
    let pt = |emu: u32| f64::from(emu) / 12_700.0;
    let line = shape.line.as_ref();
    let end = |end: Option<crate::officeart::stroke::LineEnd<'static>>| {
        end.map(|end| LineEnd {
            r#type: end.kind.into(),
            w: end.width.into(),
            len: end.length.into(),
        })
    };
    let text = shape.text.as_ref();
    let run = ShapeRun {
        inline: false,
        width_pt: placed.frame[2] / 12_700.0,
        height_pt: placed.frame[3] / 12_700.0,
        anchor_x_pt: (facts.x_emu as f64 + placed.frame[0]) / 12_700.0,
        anchor_y_pt: (facts.y_emu as f64 + placed.frame[1]) / 12_700.0,
        group_width_pt: placed.group.map(|_| facts.extent[0] as f64 / 12_700.0),
        group_height_pt: placed.group.map(|_| facts.extent[1] as f64 / 12_700.0),
        anchor_x_from_margin: matches!(facts.horizontal, "margin" | "column"),
        anchor_y_from_para: facts.vertical == "paragraph",
        anchor_x_align: facts.align[0].map(str::to_owned),
        anchor_y_align: facts.align[1].map(str::to_owned),
        anchor_x_relative_from: Some(facts.horizontal.into()),
        anchor_y_relative_from: Some(facts.vertical.into()),
        width_pct: relative_width.map(|(fraction, _)| fraction),
        height_pct: relative_height.map(|(fraction, _)| fraction),
        width_relative_from: relative_width.map(|(_, from)| from.to_owned()),
        height_relative_from: relative_height.map(|(_, from)| from.to_owned()),
        behind_doc: facts.behind,
        // Members stack in source order above the group's own layer, as
        // DOCX group members do.
        z_order: facts.z_order.saturating_add(member),
        preset_geometry: shape.preset.map(str::to_owned),
        subpaths: shape.subpaths.clone(),
        fill,
        stroke: line.map(|line| line.color.clone()),
        stroke_width: line.map_or(0.0, |line| pt(line.width_emu)),
        stroke_dash: line.and_then(|line| line.dash).map(str::to_owned),
        stroke_cap: line.map(|line| line.cap.to_owned()),
        stroke_join: line.map(|line| line.join.to_owned()),
        stroke_miter_limit: line.and_then(|line| line.miter),
        head_end: line.and_then(|line| end(line.ends[0])),
        tail_end: line.and_then(|line| end(line.ends[1])),
        rotation: placed.rotation,
        flip_h: placed.flip[0],
        flip_v: placed.flip[1],
        // MS-ODRAW 2.3.21.15 fFitShapeToText grows the shape to its text
        // (DrawingML spAutoFit). Otherwise the box is fixed and Word clips
        // overflowing text, as DrawingML noAutofit does: the Word PDF of a DOC
        // banner textbox omits the paragraph that overflows the shape.
        text_autofit: text.map(|text| if text.fit_shape { "sp" } else { "none" }.to_owned()),
        text_inset_l: text.map_or(0.0, |text| pt(text.insets[0])),
        text_inset_t: text.map_or(0.0, |text| pt(text.insets[1])),
        text_inset_r: text.map_or(0.0, |text| pt(text.insets[2])),
        text_inset_b: text.map_or(0.0, |text| pt(text.insets[3])),
        wrap_mode: Some(wrap_mode(facts.wrapping).into()),
        dist_top: pt(facts.distances[1]),
        dist_bottom: pt(facts.distances[3]),
        dist_left: pt(facts.distances[0]),
        dist_right: pt(facts.distances[2]),
        wrap_side: (facts.wrapping == 2).then(|| facts.side.into()),
        anchor_acquisition: Some(acquisition),
        ..ShapeRun::default()
    };
    let mut total = Payload(std::mem::size_of::<ShapeRun>());
    for path in &run.subpaths {
        total.add(
            std::mem::size_of::<Vec<PathCmd>>() + path.capacity() * std::mem::size_of::<PathCmd>(),
        )?;
    }
    let ends = |end: &Option<LineEnd>| {
        end.as_ref().map_or(0, |end| {
            end.r#type.capacity() + end.w.capacity() + end.len.capacity()
        })
    };
    total.add(ends(&run.head_end) + ends(&run.tail_end))?;
    if let Some(ShapeFill::Image { mime_type, .. }) = &run.fill {
        total.add(mime_type.capacity())?;
    }
    total.strings([
        run.anchor_x_align.as_ref(),
        run.anchor_y_align.as_ref(),
        run.anchor_x_relative_from.as_ref(),
        run.anchor_y_relative_from.as_ref(),
        run.width_relative_from.as_ref(),
        run.height_relative_from.as_ref(),
        run.preset_geometry.as_ref(),
        run.fill.as_ref().map(|fill| match fill {
            ShapeFill::Solid { color } => color,
            ShapeFill::Image { image_path, .. } => image_path,
            _ => unreachable!("solid or picture DOC shape fill"),
        }),
        run.stroke.as_ref(),
        run.stroke_dash.as_ref(),
        run.stroke_cap.as_ref(),
        run.stroke_join.as_ref(),
        run.text_autofit.as_ref(),
        run.wrap_mode.as_ref(),
        run.wrap_side.as_ref(),
    ])?;
    total.acquisition(run.anchor_acquisition.as_ref().expect("shape acquisition"))?;
    *remaining_bytes = remaining_bytes
        .checked_sub(total.0)
        .ok_or("OUTPUT_TOO_LARGE")?;
    Ok(DirectFloatingShape {
        shape: run,
        text: text.map(|text| text.index),
        spid,
    })
}

/// The passive BLIP reader admits PNG/JPEG rasters and validated EMF/WMF
/// metafiles (MS-ODRAW 2.2.24-25/31). They use the same extension-to-MIME
/// mapping as DOCX media parts (`image/emf`, `image/wmf`), so the shared
/// content-sniffing metafile players render them; no DOC-specific paint path.
fn mime(extension: &str) -> Result<&'static str, String> {
    match extension {
        "png" | "jpg" | "emf" | "wmf" => Ok(ooxml_common::blip::mime_from_ext(extension)),
        _ => Err(unsupported(
            "direct DOC model supports only PNG/JPEG/EMF/WMF floating pictures",
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::officeart::raster::Image;
    use std::{borrow::Cow, collections::BTreeMap};

    fn selected_store() -> Store<'static> {
        Store {
            parts: Default::default(),
            entries: Vec::new(),
            word: &[],
            table: &[],
            clx: &[],
            group_read: false,
            header_container: None,
            textboxes: [None, None],
            images: BTreeMap::from([(
                7,
                Some(Image {
                    bytes: Cow::Borrowed(b"png"),
                    extension: "png",
                }),
            )]),
            budget: 0,
            remaining_bytes: 0,
            occurrences: 0,
            selected_images: std::collections::BTreeSet::from([7]),
            omitted: false,
        }
    }

    #[test]
    fn floating_finalization_keeps_only_live_keys_and_rejects_dangling_keys() {
        let mut resources = Vec::new();
        let mut budget = 4096;
        let before = budget;
        selected_store()
            .append_referenced_direct_resources(&mut resources, &[], &mut budget)
            .unwrap();
        assert!(resources.is_empty());
        assert_eq!(budget, before);

        for key in ["legacy-doc/float/8", "legacy-doc/float/07"] {
            let mut resources = Vec::new();
            let mut budget = 4096;
            assert!(selected_store()
                .append_referenced_direct_resources(&mut resources, &[key], &mut budget)
                .is_err());
            assert!(resources.is_empty());
        }

        let mut resources = Vec::new();
        selected_store()
            .append_referenced_direct_resources(
                &mut resources,
                &["legacy-doc/float/7"],
                &mut budget,
            )
            .unwrap();
        assert_eq!(resources.len(), 1);
        assert_eq!(resources[0].key, "legacy-doc/float/7");
    }
}
