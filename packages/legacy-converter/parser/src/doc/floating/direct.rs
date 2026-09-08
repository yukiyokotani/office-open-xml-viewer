//! Direct-model projection of already validated main-story floating pictures.

use super::{unsupported, ResolvedDrawing, Store};
use crate::doc::pictures::DirectPictureResource;
use docx_model::{
    AnchorAcquisitionWire, AnchorAxisChoiceWire, AnchorAxisWire, AnchorBehaviorWire,
    AnchorEdgesWire, AnchorExtentWire, AnchorSimplePositionWire, AnchorValueStatusWire,
    AnchorWrapKindWire, AnchorWrapWire, ImageRun,
};

#[derive(Debug)]
pub(in crate::doc) struct DirectFloatingPicture {
    pub image: ImageRun,
    pub occurrence_id: String,
}

impl Store<'_> {
    pub(in crate::doc) fn direct_picture(
        &mut self,
        cp: usize,
        remaining_bytes: &mut usize,
    ) -> Result<Option<DirectFloatingPicture>, String> {
        let Some(facts) = self.resolve(cp)? else {
            return Ok(None);
        };
        let mime_type = mime(facts.extension)?.to_string();
        let image_path = format!("legacy-doc/float/{}", facts.image_index);
        let occurrence_id = format!("legacy-doc-float-{}", facts.occurrence);
        let acquisition = acquisition(&facts, occurrence_id.clone());
        let [top, bottom, left, right] = facts.crop;
        let image = ImageRun {
            image_path,
            mime_type,
            svg_image_path: None,
            src_rect: (facts.crop != [0; 4]).then_some(ooxml_common::blip::SrcRect {
                l: left as f64 / 100_000.0,
                t: top as f64 / 100_000.0,
                r: right as f64 / 100_000.0,
                b: bottom as f64 / 100_000.0,
            }),
            width_pt: facts.extent[0] as f64 / 12_700.0,
            height_pt: facts.extent[1] as f64 / 12_700.0,
            rotation: 0.0,
            flip_h: facts.flip[0],
            flip_v: facts.flip[1],
            anchor: true,
            anchor_x_pt: facts.x_emu as f64 / 12_700.0,
            anchor_y_pt: facts.y_emu as f64 / 12_700.0,
            anchor_x_from_margin: matches!(facts.horizontal, "margin" | "column"),
            anchor_y_from_para: facts.vertical == "paragraph",
            color_replace_from: None,
            duotone: None,
            alpha: None,
            wrap_mode: Some(
                match facts.wrapping {
                    1 => "topAndBottom",
                    2 => "square",
                    3 => "none",
                    _ => unreachable!(),
                }
                .into(),
            ),
            dist_top: facts.distances[1] as f64 / 12_700.0,
            dist_bottom: facts.distances[3] as f64 / 12_700.0,
            dist_left: facts.distances[0] as f64 / 12_700.0,
            dist_right: facts.distances[2] as f64 / 12_700.0,
            wrap_side: (facts.wrapping == 2).then(|| facts.side.into()),
            allow_overlap: facts.overlap,
            anchor_x_align: None,
            anchor_y_align: None,
            anchor_x_relative_from: Some(facts.horizontal.into()),
            anchor_y_relative_from: Some(facts.vertical.into()),
            anchor_acquisition: Some(acquisition),
        };
        let required = image_payload(&image, &occurrence_id)?;
        *remaining_bytes = remaining_bytes
            .checked_sub(required)
            .ok_or("OUTPUT_TOO_LARGE")?;
        self.selected_images.insert(facts.image_index);
        Ok(Some(DirectFloatingPicture {
            occurrence_id,
            image,
        }))
    }

    pub(in crate::doc) fn append_direct_resources(
        self,
        resources: &mut Vec<DirectPictureResource>,
        remaining_bytes: &mut usize,
    ) -> Result<(), String> {
        let count = self.selected_images.len();
        let old_capacity = resources.capacity();
        let minimum_capacity = resources
            .len()
            .checked_add(count)
            .ok_or("OUTPUT_TOO_LARGE")?;
        let minimum_added = minimum_capacity.saturating_sub(old_capacity);
        let mut required = minimum_added
            .checked_mul(std::mem::size_of::<DirectPictureResource>())
            .ok_or("OUTPUT_TOO_LARGE")?;
        for index in &self.selected_images {
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
            if !selected_images.contains(&index) {
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

fn acquisition(f: &ResolvedDrawing, occurrence_id: String) -> AnchorAcquisitionWire {
    let axis = |relative: &'static str, value: i64| AnchorAxisWire {
        relative_from: Some(relative.into()),
        relative_from_status: AnchorValueStatusWire::Valid,
        choice: AnchorAxisChoiceWire::Offset {
            value_pt: value as f64 / 12_700.0,
        },
    };
    AnchorAcquisitionWire {
        occurrence_id,
        simple_position: AnchorSimplePositionWire {
            enabled: Some(false),
            status: AnchorValueStatusWire::Valid,
            x_pt: Some(0.0),
            x_status: AnchorValueStatusWire::Valid,
            y_pt: Some(0.0),
            y_status: AnchorValueStatusWire::Valid,
        },
        horizontal: axis(f.horizontal, f.x_emu),
        vertical: axis(f.vertical, f.y_emu),
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
                3 => AnchorWrapKindWire::None,
                _ => unreachable!(),
            },
            authored_kinds: vec![match f.wrapping {
                1 => "wrapTopAndBottom",
                2 => "wrapSquare",
                3 => "wrapNone",
                _ => unreachable!(),
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
        ..AnchorAcquisitionWire::default()
    }
}

fn image_payload(image: &ImageRun, host_occurrence_id: &String) -> Result<usize, String> {
    let mut total = std::mem::size_of::<ImageRun>();
    let mut add = |bytes: usize| -> Result<(), String> {
        total = total.checked_add(bytes).ok_or("OUTPUT_TOO_LARGE")?;
        Ok(())
    };
    for value in [
        Some(&image.image_path),
        Some(&image.mime_type),
        image.wrap_mode.as_ref(),
        image.wrap_side.as_ref(),
        image.anchor_x_relative_from.as_ref(),
        image.anchor_y_relative_from.as_ref(),
        Some(host_occurrence_id),
    ]
    .into_iter()
    .flatten()
    {
        add(value.capacity())?;
    }
    let facts = image
        .anchor_acquisition
        .as_ref()
        .expect("floating acquisition");
    add(facts.occurrence_id.capacity())?;
    for value in [
        facts.horizontal.relative_from.as_ref(),
        facts.vertical.relative_from.as_ref(),
        facts.wrap.side.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        add(value.capacity())?;
    }
    add(facts
        .wrap
        .authored_kinds
        .capacity()
        .checked_mul(std::mem::size_of::<String>())
        .ok_or("OUTPUT_TOO_LARGE")?)?;
    for value in &facts.wrap.authored_kinds {
        add(value.capacity())?;
    }
    Ok(total)
}

fn mime(extension: &str) -> Result<&'static str, String> {
    match extension {
        "png" | "jpg" => Ok(ooxml_common::blip::mime_from_ext(extension)),
        _ => Err(unsupported(
            "direct DOC model supports only PNG/JPEG floating pictures",
        )),
    }
}
