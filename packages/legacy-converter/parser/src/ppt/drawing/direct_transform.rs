//! Project already-decoded OfficeArt placement into the receiving model.
//! Coordinates are EMU at this boundary; MS-ODRAW rotation has already been
//! converted to DrawingML's 1/60000 degree units by PropertiesStorage.
use super::*;
use pptx_model::{GroupTransform, Transform};

pub(super) fn leaf<R, C, S>(shape: &ShapeStorage<R, C, S>) -> Result<Transform, String> {
    let anchor = shape
        .anchor
        .ok_or_else(|| unsupported("missing PowerPoint shape anchor"))?;
    Ok(at(shape, anchor))
}

/// OfficeArt placement -> DrawingML placement.
///
/// MS-ODRAW 2.3.18.5 only defines the angle (clockwise, about the centre).
/// Two further rules are Office behaviour, established by a PowerPoint
/// control (`converter-rotation-matrix-v1`: one custom polygon at 0, 30,
/// 44.99, 45, 45.01, 60, 90, 120, 134.99, 135 ... 359 degrees, each with all
/// four flip combinations, authored as PPTX and saved as PPT by PowerPoint;
/// all 120 PPTX transforms are reproduced by these rules) and confirmed on
/// grouped children and a rotated group in two corpus decks:
/// - The stored angle is negated when exactly one of fFlipH/fFlipV is set;
///   DrawingML applies the flip first and then the unnegated rotation.
/// - When the stored angle, normalised to [0, 360), lies in [45, 135) or
///   [225, 315), the anchor holds the rotated bounds: width and height are
///   swapped about the same centre. The half-open intervals are exact at the
///   44.99/45 and 134.99/135 boundaries. A group's child coordinate space is
///   never swapped.
fn at<R, C, S>(shape: &ShapeStorage<R, C, S>, anchor: Rect) -> Transform {
    let flip_h = shape.flags & 0x40 != 0;
    let flip_v = shape.flags & 0x80 != 0;
    let stored = shape.props.rotation;
    let normalized = stored.rem_euclid(21_600_000);
    let swapped = (2_700_000..8_100_000).contains(&normalized)
        || (13_500_000..18_900_000).contains(&normalized);
    let (x, y, cx, cy) = if swapped {
        (
            anchor.x + (anchor.w - anchor.h) / 2,
            anchor.y + (anchor.h - anchor.w) / 2,
            anchor.h,
            anchor.w,
        )
    } else {
        (anchor.x, anchor.y, anchor.w, anchor.h)
    };
    let rotation = if flip_h != flip_v { -stored } else { stored };
    Transform {
        x,
        y,
        cx,
        cy,
        rot: rotation as f64 / 60000.0,
        flip_h,
        flip_v,
    }
}

/// A top-level patriarch defines no transform. Ordinary groups retain the
/// exact anchor/child coordinate spaces consumed by the shared model helper.
pub(super) fn group<R, C, S>(
    shape: &ShapeStorage<R, C, S>,
    nested: bool,
) -> Result<Option<GroupTransform>, String> {
    if shape.flags & 1 == 0 {
        return Err(unsupported("missing PowerPoint group flag"));
    }
    if shape.flags & 4 != 0 {
        return if nested {
            Err(unsupported("nested PowerPoint patriarch group"))
        } else {
            Ok(None)
        };
    }
    let anchor = shape
        .anchor
        .ok_or_else(|| unsupported("missing PowerPoint group anchor"))?;
    let child = shape
        .child_space
        .filter(|r| r.w > 0 && r.h > 0)
        .ok_or_else(|| unsupported("invalid PowerPoint group coordinate space"))?;
    let t = at(shape, anchor);
    Ok(Some(GroupTransform {
        x: t.x,
        y: t.y,
        cx: t.cx,
        cy: t.cy,
        ch_x: child.x,
        ch_y: child.y,
        ch_cx: child.w,
        ch_cy: child.h,
        rot: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
    }))
}

/// Ancestors are ordered outermost first. Applying in reverse preserves the
/// same local-to-parent composition as nested DrawingML groups, without XML.
pub(super) fn flatten(mut leaf: Transform, ancestors: &[GroupTransform]) -> Transform {
    for group in ancestors.iter().rev() {
        leaf = group.apply_to_transform(leaf);
    }
    leaf
}

#[cfg(test)]
mod tests {
    use super::*;

    fn shape(flags: u32) -> Shape<'static> {
        ShapeStorage {
            id: 1,
            kind: 1,
            flags,
            anchor: Some(Rect {
                x: -100,
                y: 200,
                w: 600,
                h: 400,
            }),
            child_space: Some(Rect {
                x: 10,
                y: -20,
                w: 300,
                h: 100,
            }),
            textbox: None,
            style9: None,
            placeholder: None,
            props: Properties::default(),
        }
    }

    #[test]
    fn negates_rotation_under_a_single_flip_and_keeps_all_flip_combinations() {
        for (flags, rot) in [(0, -22.5), (0x40, 22.5), (0x80, 22.5), (0xc0, -22.5)] {
            let mut shape = shape(flags);
            shape.props.rotation = -1350000;
            let t = leaf(&shape).unwrap();
            assert_eq!((t.x, t.y, t.cx, t.cy), (-100, 200, 600, 400));
            assert_eq!(t.rot, rot);
            assert_eq!(t.flip_h, flags & 0x40 != 0);
            assert_eq!(t.flip_v, flags & 0x80 != 0);
        }
    }

    #[test]
    fn swaps_rotated_bounds_on_the_observed_half_open_intervals() {
        // Anchor 600 x 400 centred at (200, 400).
        for (degrees_60k, swapped) in [
            (2_699_400, false),  // 44.99
            (2_700_000, true),   // 45
            (8_099_400, true),   // 134.99
            (8_100_000, false),  // 135
            (13_500_000, true),  // 225
            (18_900_000, false), // 315
            (-2_700_000, false), // -45 == 315
            (-2_700_600, true),  // -45.01 == 314.99
        ] {
            let mut shape = shape(0);
            shape.props.rotation = degrees_60k;
            let t = leaf(&shape).unwrap();
            let expected = if swapped {
                (0, 100, 400, 600)
            } else {
                (-100, 200, 600, 400)
            };
            assert_eq!((t.x, t.y, t.cx, t.cy), expected, "{degrees_60k}");
        }
        // The swap follows the stored angle, before flip negation: a stored
        // -45 with one flip keeps its bounds and renders at +45.
        let mut shape = shape(0x40);
        shape.props.rotation = -2_700_000;
        let t = leaf(&shape).unwrap();
        assert_eq!((t.cx, t.cy, t.rot), (600, 400, 45.0));
    }

    #[test]
    fn rotated_group_swaps_its_extent_but_not_its_child_space() {
        let mut s = shape(1);
        s.props.rotation = -5_400_000;
        let g = group(&s, false).unwrap().unwrap();
        assert_eq!((g.x, g.y, g.cx, g.cy), (0, 100, 400, 600));
        assert_eq!((g.ch_x, g.ch_y, g.ch_cx, g.ch_cy), (10, -20, 300, 100));
        assert_eq!(g.rot, -90.0);
    }

    #[test]
    fn retains_group_units_and_composes_inner_before_outer() {
        let g = group(&shape(1), false).unwrap().unwrap();
        let t = Transform {
            x: 10,
            y: -20,
            cx: 150,
            cy: 50,
            ..Transform::default()
        };
        let mapped = flatten(t, &[g]);
        assert_eq!(
            (mapped.x, mapped.y, mapped.cx, mapped.cy),
            (-100, 200, 300, 200)
        );
        let outer = GroupTransform {
            x: 100,
            y: 200,
            cx: 200,
            cy: 300,
            ch_cx: 100,
            ch_cy: 100,
            ..GroupTransform::default()
        };
        let inner = GroupTransform {
            x: 10,
            y: 20,
            cx: 100,
            cy: 100,
            ch_cx: 100,
            ch_cy: 100,
            ..GroupTransform::default()
        };
        let mapped = flatten(
            Transform {
                x: 1,
                y: 2,
                cx: 3,
                cy: 4,
                ..Transform::default()
            },
            &[outer, inner],
        );
        assert_eq!(
            (mapped.x, mapped.y, mapped.cx, mapped.cy),
            (122, 266, 6, 12)
        );
    }

    #[test]
    fn rejects_missing_or_degenerate_spaces_but_not_zero_sized_leaf() {
        let mut s = shape(1);
        s.anchor = None;
        assert!(leaf(&s).is_err());
        assert!(group(&s, false).is_err());
        let mut s = shape(1);
        s.child_space.as_mut().unwrap().w = 0;
        assert!(group(&s, false).is_err());
        s.anchor.as_mut().unwrap().w = 0;
        assert_eq!(leaf(&s).unwrap().cx, 0);
        assert!(group(&shape(0), false).is_err());
        assert!(group(&shape(5), true).is_err());
        let mut patriarch = shape(5);
        patriarch.anchor = None;
        patriarch.child_space = None;
        assert!(group(&patriarch, false).unwrap().is_none());
    }
}
