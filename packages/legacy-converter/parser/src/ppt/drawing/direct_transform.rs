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

fn at<R, C, S>(shape: &ShapeStorage<R, C, S>, anchor: Rect) -> Transform {
    Transform {
        x: anchor.x,
        y: anchor.y,
        cx: anchor.w,
        cy: anchor.h,
        rot: shape.props.rotation as f64 / 60000.0,
        flip_h: shape.flags & 0x40 != 0,
        flip_v: shape.flags & 0x80 != 0,
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
    fn preserves_signed_placement_rotation_and_all_flip_combinations() {
        for flags in [0, 0x40, 0x80, 0xc0] {
            let mut shape = shape(flags);
            shape.props.rotation = -1350000;
            let t = leaf(&shape).unwrap();
            assert_eq!((t.x, t.y, t.cx, t.cy), (-100, 200, 600, 400));
            assert_eq!(t.rot, -22.5);
            assert_eq!(t.flip_h, flags & 0x40 != 0);
            assert_eq!(t.flip_v, flags & 0x80 != 0);
        }
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
