//! Isolated Word compatibility rules for DrawingML group transforms.

use ooxml_common::drawing::{DrawingGroupTransform, DrawingRect};

fn is_exact_odd_quarter_turn(rotation_60000: i64) -> bool {
    rotation_60000.rem_euclid(5_400_000) == 0
        && rotation_60000.div_euclid(5_400_000).rem_euclid(2) == 1
}

pub(crate) fn word_group_requires_hierarchy_compatibility(
    transform: DrawingGroupTransform,
    rotation_60000: i64,
) -> bool {
    transform.non_neutral_group_levels() > 1
        && is_exact_odd_quarter_turn(rotation_60000)
        && transform.scale_x != transform.scale_y
}

/// Apply Word's exact-quarter-turn scale order for a leaf under one effective
/// (non-neutral) group transform.
///
/// ECMA-376 Part 1 Annex L §L.4.7.4–§L.4.7.5 applies conventional authored-axis
/// scale before the leaf rotation. This is an explicitly scoped Word
/// compatibility override of that pipeline.
/// [MS-OE376] §2.1.1360 defines the Office group ratio as `ext / chExt`.
/// Word-produced reference output additionally shows that a directly grouped
/// leaf at an exact odd quarter turn uses that ratio on post-rotation page
/// axes. Neutral translation wrappers are transparent; multiple effective group
/// transforms still require a hierarchy-aware retained transform and remain on
/// the Annex L fallback pending broader Word evidence.
pub(crate) fn apply_word_direct_group_rect(
    transform: DrawingGroupTransform,
    rect: DrawingRect,
    rotation_60000: i64,
) -> DrawingRect {
    debug_assert!(
        (rect.rotation_degrees - rotation_60000 as f64 / 60_000.0).abs() < 1e-9,
        "authored rotation units and derived degrees must describe the same transform",
    );
    let mapped = transform.apply_rect(rect);
    if transform.non_neutral_group_levels() != 1 || !is_exact_odd_quarter_turn(rotation_60000) {
        return mapped;
    }

    let center_x = mapped.x + mapped.width / 2.0;
    let center_y = mapped.y + mapped.height / 2.0;
    let width = rect.width * transform.scale_y;
    let height = rect.height * transform.scale_x;
    DrawingRect {
        x: center_x - width / 2.0,
        y: center_y - height / 2.0,
        width,
        height,
        ..mapped
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use ooxml_common::drawing::DrawingGroupSpec;

    fn group(rotation_degrees: f64) -> DrawingGroupSpec {
        DrawingGroupSpec {
            off_x: 0.0,
            off_y: 0.0,
            ext_x: 127_000.0,
            ext_y: 254_000.0,
            child_off_x: 0.0,
            child_off_y: 0.0,
            child_ext_x: 127_000.0,
            child_ext_y: 127_000.0,
            rotation_degrees,
            flip_h: false,
            flip_v: false,
        }
    }

    fn leaf(rotation_degrees: f64) -> DrawingRect {
        DrawingRect {
            x: 0.0,
            y: 50_800.0,
            width: 127_000.0,
            height: 25_400.0,
            rotation_degrees,
            flip_h: false,
            flip_v: false,
        }
    }

    fn identity_group() -> DrawingGroupSpec {
        DrawingGroupSpec {
            off_x: 0.0,
            off_y: 0.0,
            ext_x: 127_000.0,
            ext_y: 127_000.0,
            child_off_x: 0.0,
            child_off_y: 0.0,
            child_ext_x: 127_000.0,
            child_ext_y: 127_000.0,
            rotation_degrees: 0.0,
            flip_h: false,
            flip_v: false,
        }
    }

    #[test]
    fn exchanges_axes_for_a_direct_exact_quarter_turn() {
        let mapped = apply_word_direct_group_rect(
            DrawingGroupTransform::from_group(group(0.0)),
            leaf(90.0),
            5_400_000,
        );
        assert!((mapped.x + 63_500.0).abs() < 1e-6);
        assert!((mapped.y - 114_300.0).abs() < 1e-6);
        assert!((mapped.width - 254_000.0).abs() < 1e-6);
        assert!((mapped.height - 25_400.0).abs() < 1e-6);
    }

    #[test]
    fn identity_wrapper_does_not_change_quarter_turn_compatibility() {
        let direct = DrawingGroupTransform::from_group(group(0.0));
        let wrapped = DrawingGroupTransform::from_group(identity_group()).compose_group(group(0.0));

        assert_eq!(
            apply_word_direct_group_rect(wrapped, leaf(90.0), 5_400_000),
            apply_word_direct_group_rect(direct, leaf(90.0), 5_400_000),
        );
    }

    #[test]
    fn translation_only_wrapper_keeps_quarter_turn_dimensions_eligible() {
        let direct = DrawingGroupTransform::from_group(group(0.0));
        let mut translation = identity_group();
        translation.off_x = 63_500.0;
        translation.off_y = 25_400.0;
        let wrapped = DrawingGroupTransform::from_group(translation).compose_group(group(0.0));

        let direct_rect = apply_word_direct_group_rect(direct, leaf(90.0), 5_400_000);
        let wrapped_rect = apply_word_direct_group_rect(wrapped, leaf(90.0), 5_400_000);
        assert_eq!(wrapped.non_neutral_group_levels(), 1);
        assert_eq!(wrapped_rect.width, direct_rect.width);
        assert_eq!(wrapped_rect.height, direct_rect.height);
    }

    #[test]
    fn near_quarter_turn_stays_on_the_annex_l_path() {
        let transform = DrawingGroupTransform::from_group(group(0.0));
        let near_quarter = leaf(90.0 - 1.0 / 60_000.0);
        assert_eq!(
            apply_word_direct_group_rect(transform, near_quarter, 5_399_999),
            transform.apply_rect(near_quarter),
        );
    }

    #[test]
    fn leaves_nested_rotated_groups_on_the_annex_l_path() {
        let nested = DrawingGroupTransform::from_group(group(90.0)).compose_group(group(0.0));
        let mapped = apply_word_direct_group_rect(nested, leaf(0.0), 0);
        let normative = nested.apply_rect(leaf(0.0));
        assert_eq!(mapped, normative);
        assert_eq!(nested.non_neutral_group_levels(), 2);
    }
}
