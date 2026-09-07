//! Synthetic end-to-end coverage for retained classic OfficeArt gradients.

use super::*;
use pptx_model::{Fill, SlideElement};
use std::rc::Rc;

fn record(options: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    [
        options.to_le_bytes().as_slice(),
        kind.to_le_bytes().as_slice(),
        (payload.len() as u32).to_le_bytes().as_slice(),
        payload,
    ]
    .concat()
}

fn shade_array() -> Vec<u8> {
    [
        2u16.to_le_bytes().as_slice(),
        2u16.to_le_bytes().as_slice(),
        8u16.to_le_bytes().as_slice(),
        0x0000_00ffu32.to_le_bytes().as_slice(),
        0u32.to_le_bytes().as_slice(),
        0x00ff_0000u32.to_le_bytes().as_slice(),
        65_536u32.to_le_bytes().as_slice(),
    ]
    .concat()
}

fn gradient_properties(extra: &[(u16, u32)], gradient_scalar: Option<u32>) -> Vec<u8> {
    let shade = shade_array();
    let mut values = vec![
        (0x180u16, 4u32),
        (0x181, 0x0000_00ff),
        (0x183, 0x00ff_0000),
        (0x18b, 0),
        (0x18c, 100),
    ];
    values.extend_from_slice(extra);
    let gradient = match gradient_scalar {
        Some(value) => (0x0197u16, value, false),
        None => (0x8197u16, shade.len() as u32, true),
    };
    let mut entries: Vec<_> = values
        .iter()
        .copied()
        .map(|(id, value)| (id, value, false))
        .chain(std::iter::once(gradient))
        .collect();
    entries.sort_unstable_by_key(|(id, _, _)| id & 0x3fff);
    let mut body = Vec::new();
    for (id, value, _) in entries.iter().copied() {
        body.extend(id.to_le_bytes());
        body.extend(value.to_le_bytes());
    }
    if entries.iter().any(|(_, _, complex)| *complex) {
        body.extend(shade);
    }
    record(((entries.len() as u16) << 4) | 3, 0xf00b, &body)
}

fn scalar_properties(values: &[(u16, u32)]) -> Vec<u8> {
    let mut body = Vec::new();
    for (id, value) in values.iter().copied() {
        body.extend(id.to_le_bytes());
        body.extend(value.to_le_bytes());
    }
    record(((values.len() as u16) << 4) | 3, 0xf00b, &body)
}

fn shape(flags: u32, properties: Vec<u8>) -> Vec<u8> {
    record(
        15,
        0xf004,
        &[
            record(
                (1 << 4) | 2,
                0xf00a,
                &[42u32.to_le_bytes(), flags.to_le_bytes()].concat(),
            ),
            record(
                0,
                0xf010,
                &[0i16, 0, 576, 576]
                    .into_iter()
                    .flat_map(i16::to_le_bytes)
                    .collect::<Vec<_>>(),
            ),
            properties,
        ]
        .concat(),
    )
}

fn nested_shape(properties: Vec<u8>) -> Vec<u8> {
    record(
        15,
        0xf004,
        &[
            record(
                (1 << 4) | 2,
                0xf00a,
                &[43u32.to_le_bytes(), 0xa00u32.to_le_bytes()].concat(),
            ),
            record(
                0,
                0xf00f,
                &[0i32, 0, 288, 288]
                    .into_iter()
                    .flat_map(i32::to_le_bytes)
                    .collect::<Vec<_>>(),
            ),
            properties,
        ]
        .concat(),
    )
}

fn group(group_flags: u32, rotation: Option<u32>, child: Vec<u8>) -> Vec<u8> {
    let mut head = vec![
        record(
            2,
            0xf00a,
            &[44u32.to_le_bytes(), (1 | group_flags).to_le_bytes()].concat(),
        ),
        record(
            0,
            0xf010,
            &[0i32, 0, 576, 576]
                .into_iter()
                .flat_map(i32::to_le_bytes)
                .collect::<Vec<_>>(),
        ),
        record(
            1,
            0xf009,
            &[0i32, 0, 288, 288]
                .into_iter()
                .flat_map(i32::to_le_bytes)
                .collect::<Vec<_>>(),
        ),
    ];
    if let Some(rotation) = rotation {
        head.push(scalar_properties(&[(4, rotation)]));
    }
    record(
        15,
        0xf003,
        &[record(15, 0xf004, &head.concat()), child].concat(),
    )
}

fn drawing(shapes: &[Vec<u8>]) -> Vec<u8> {
    record(15, 1036, &record(15, 0xf002, &shapes.concat()))
}

fn slide_container(drawing: &[u8]) -> Vec<u8> {
    record(15, SLIDE_CONTAINER, drawing)
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
        object_masters: vec![Rc::from([])],
        size: (720, 540),
    }
}

fn native(
    backing: &[u8],
    presentation: persist::OwnedPresentation,
) -> Result<pptx_model::Slide, String> {
    let mut media = media::SpanStore::new(Vec::new());
    direct_model::slide(
        0,
        &presentation,
        backing,
        None,
        &mut media,
        &mut 10_000,
        &mut 10_000,
        &mut 1_000_000,
    )
}

fn assert_gradient(fill: &Option<Fill>) {
    let Some(Fill::Gradient {
        stops,
        angle,
        grad_type,
        rot_with_shape,
        ..
    }) = fill
    else {
        panic!("expected linear gradient")
    };
    assert_eq!(grad_type, "linear");
    assert_eq!(*angle, 90.0);
    assert_eq!(*rot_with_shape, Some(false));
    assert_eq!(stops.len(), 2);
    assert_eq!(
        (stops[0].position, stops[0].color.as_str()),
        (0.0, "FF0000")
    );
    assert_eq!(
        (stops[1].position, stops[1].color.as_str()),
        (1.0, "0000FF")
    );
}

#[test]
fn foreground_xml_and_native_keep_quantized_stops_for_all_leaf_flips() {
    for flip in [0, 0x40, 0x80, 0xc0] {
        let tree = drawing(&[shape(0xa00 | flip, gradient_properties(&[], None))]);
        let xml = render(
            &tree,
            &[],
            &mut 10_000,
            &mut 10_000,
            &mut 1_000_000,
            None,
            None,
        )
        .unwrap()
        .unwrap();
        assert!(xml.contains("<a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/>"));
        assert!(xml.contains("<a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/>"));
        assert!(xml.contains("<a:lin ang=\"5400000\"/>"));

        let document = slide_container(&tree);
        let span = record_span_with_end(&document, 0, &mut 100, "gradient slide")
            .unwrap()
            .0;
        let model = native(&document, presentation(span)).unwrap();
        let SlideElement::Shape(shape) = &model.elements[0] else {
            panic!("expected shape")
        };
        assert_eq!(
            (shape.flip_h, shape.flip_v),
            (flip & 0x40 != 0, flip & 0x80 != 0)
        );
        assert_gradient(&shape.fill);
    }
}

#[test]
fn master_gradient_inherits_but_local_scalar_zero_resets_it() {
    for reset in [false, true] {
        let local_properties = if reset {
            gradient_properties(&[(0x301, 1)], Some(0))
        } else {
            scalar_properties(&[(0x301, 1)])
        };
        let local = drawing(&[shape(0xa20, local_properties)]);
        let document = slide_container(&local);
        let span = record_span_with_end(&document, 0, &mut 100, "gradient slide")
            .unwrap()
            .0;
        let mut p = presentation(span);
        let shade = shade_array();
        let shade_span = ByteSpan::new(0..shade.len(), shade.len(), "master gradient").unwrap();
        let mut paint = paint::Paint::default();
        for (id, value) in [
            (0x180, 4),
            (0x181, 0x0000_00ff),
            (0x183, 0x00ff_0000),
            (0x18b, 0),
            (0x18c, 100),
        ] {
            paint.property(id, value).unwrap();
        }
        let mut gradient = crate::officeart::gradient::Spanned::default();
        gradient.set(shade_span);
        p.shape_masters
            .insert(shape_master::Node {
                id: 1,
                parent: None,
                text_type: None,
                direct: Vec::new(),
                base: None,
                paint,
                geometry: crate::officeart::geometry::SpannedGeometry::default(),
                gradient,
            })
            .unwrap();
        p.shape_masters.finish(&mut 100).unwrap();

        // The master shade span deliberately points at a separately owned
        // backing, so combine it with the slide and rebuild the absolute span.
        let combined = [shade, document].concat();
        let slide_span =
            record_span_with_end(&combined, shade_array().len(), &mut 100, "gradient slide")
                .unwrap()
                .0;
        p.slides[0].0 = slide_span;
        let empty_styles = Vec::new();
        let empty_types = Vec::new();
        let empty_numbers = Vec::new();
        let xml = render(
            &local,
            &[],
            &mut 10_000,
            &mut 10_000,
            &mut 1_000_000,
            Some(TextContext {
                fonts: &[],
                styles: &empty_styles,
                scheme: None,
                types: &empty_types,
                master: None,
                shapes: Some(&p.shape_masters),
                backing: &combined,
                outline_slide_numbers: &empty_numbers,
                slide_number: 1,
            }),
            None,
        )
        .unwrap()
        .unwrap();
        assert_eq!(xml.contains("<a:gradFill"), !reset);
        let model = native(&combined, p);
        if reset {
            let model = model.unwrap();
            let SlideElement::Shape(shape) = &model.elements[0] else {
                panic!("shape")
            };
            assert!(!matches!(shape.fill, Some(Fill::Gradient { .. })));
        } else {
            let model = model.unwrap();
            let SlideElement::Shape(shape) = &model.elements[0] else {
                panic!("shape")
            };
            assert_gradient(&shape.fill);
        }
    }
}

#[test]
fn native_background_uses_the_retained_gradient_span() {
    let shade = shade_array();
    let tree = drawing(&[]);
    let document = slide_container(&tree);
    let combined = [shade.clone(), document].concat();
    let span = record_span_with_end(&combined, shade.len(), &mut 100, "gradient slide")
        .unwrap()
        .0;
    let mut p = presentation(span);
    let mut paint = paint::Paint::default();
    for (id, value) in [
        (0x180, 4),
        (0x181, 0x0000_00ff),
        (0x183, 0x00ff_0000),
        (0x18b, 0),
        (0x18c, 100),
    ] {
        paint.property(id, value).unwrap();
    }
    let mut gradient = crate::officeart::gradient::Spanned::default();
    gradient.set(ByteSpan::new(0..shade.len(), combined.len(), "background gradient").unwrap());
    p.backgrounds[0] = Some(SpannedBackground { paint, gradient });
    let model = native(&combined, p).unwrap();
    assert_gradient(&model.background);

    let mut borrowed = crate::officeart::gradient::Borrowed::default();
    borrowed.set(&shade);
    let mut media = media::Store::new(&[], &[]);
    let xml =
        super::super::background_xml(&paint, &borrowed, None, &mut media, &mut 10_000, 1_000_000)
            .unwrap();
    assert!(xml.contains("<p:bg><p:bgPr><a:gradFill"));
    assert!(xml.contains("<a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/>"));
    assert!(xml.contains("<a:lin ang=\"5400000\"/>"));
}

#[test]
fn nonzero_leaf_rotation_vetoes_gradient_projection() {
    let tree = drawing(&[shape(0xa00, gradient_properties(&[(4, 45 << 16)], None))]);
    let xml = render(
        &tree,
        &[],
        &mut 10_000,
        &mut 10_000,
        &mut 1_000_000,
        None,
        None,
    )
    .unwrap()
    .unwrap();
    assert!(!xml.contains("<a:gradFill"));
    let document = slide_container(&tree);
    let span = record_span_with_end(&document, 0, &mut 100, "gradient slide")
        .unwrap()
        .0;
    let model = native(&document, presentation(span)).unwrap();
    let SlideElement::Shape(shape) = &model.elements[0] else {
        panic!("shape")
    };
    assert!(!matches!(shape.fill, Some(Fill::Gradient { .. })));
}

#[test]
fn scaled_groups_admit_but_rotated_or_reflected_ancestors_veto_gradients() {
    for (group_flags, rotation, admitted) in [
        (0, None, true),
        (0, Some(45 << 16), false),
        (0x40, None, false),
        (0x80, None, false),
    ] {
        let grouped = group(
            group_flags,
            rotation,
            nested_shape(gradient_properties(&[], None)),
        );
        let tree = drawing(&[grouped]);
        let xml = render(
            &tree,
            &[],
            &mut 10_000,
            &mut 10_000,
            &mut 1_000_000,
            None,
            None,
        )
        .unwrap()
        .unwrap();
        assert_eq!(xml.contains("<a:gradFill"), admitted);

        let document = slide_container(&tree);
        let span = record_span_with_end(&document, 0, &mut 100, "gradient slide")
            .unwrap()
            .0;
        let model = native(&document, presentation(span)).unwrap();
        let SlideElement::Shape(shape) = &model.elements[0] else {
            panic!("shape")
        };
        assert_eq!(matches!(shape.fill, Some(Fill::Gradient { .. })), admitted);
    }
}

#[test]
fn malformed_gradient_and_outer_short_budgets_fail_without_partial_admission() {
    let malformed = gradient_properties(&[], None);
    let mut truncated = malformed.clone();
    truncated.pop();
    let tree = drawing(&[shape(0xa00, truncated)]);
    assert!(render(
        &tree,
        &[],
        &mut 10_000,
        &mut 10_000,
        &mut 1_000_000,
        None,
        None
    )
    .is_err());

    let tree = drawing(&[shape(0xa00, gradient_properties(&[], None))]);
    assert!(render(&tree, &[], &mut 1, &mut 10_000, &mut 1_000_000, None, None).is_err());
    assert!(render(&tree, &[], &mut 10_000, &mut 10_000, &mut 1, None, None).is_err());
}
