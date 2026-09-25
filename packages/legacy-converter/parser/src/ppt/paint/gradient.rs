//! Adapter from retained OfficeArt shade facts to presentation-model gradients.

use super::Paint;
use crate::officeart::gradient as office_gradient;
use crate::ppt::{scheme, unsupported};
use pptx_model::{Fill, GradStop};
use std::fmt::Write;
use std::mem::size_of;

pub(in crate::ppt) struct ResolvedGradient {
    projection: office_gradient::projection::Projection,
    /// Per-stop 16.16 opacity, parallel to `projection.stops`.
    alphas: Vec<u32>,
    scaled: bool,
    rotate_with_shape: bool,
}

impl Paint {
    /// PowerPoint's adapter over the shared linear-shade rules
    /// (`officeart::paint::Paint::linear_shade`), resolving slide scheme
    /// colours for every projected stop.
    pub(in crate::ppt) fn project_gradient(
        &self,
        source: &office_gradient::Borrowed<'_>,
        allow_fill: bool,
        colors: Option<&scheme::Scheme>,
        work_budget: &mut usize,
        byte_budget: &mut usize,
    ) -> Result<Option<ResolvedGradient>, String> {
        let Some(shade) = self.linear_shade(source, allow_fill, work_budget, byte_budget)? else {
            return Ok(None);
        };
        let mut projection = shade.projection;
        for stop in &mut projection.stops {
            let Some(color) = scheme::drawing(stop.color, colors) else {
                return Err(unsupported("PowerPoint gradient colour cannot be resolved"));
            };
            stop.color = color;
        }
        Ok(Some(ResolvedGradient {
            projection,
            alphas: shade.alphas,
            scaled: shade.scaled,
            rotate_with_shape: shade.rotate_with_shape,
        }))
    }
}

impl ResolvedGradient {
    pub(in crate::ppt) fn into_model(self, byte_budget: &mut usize) -> Result<Fill, String> {
        let count = self.projection.stops.len();
        let bytes = count
            .checked_mul(size_of::<GradStop>() + 8)
            .and_then(|value| value.checked_add("linear".len()))
            .ok_or_else(|| unsupported("PowerPoint gradient model budget overflow"))?;
        *byte_budget = byte_budget
            .checked_sub(bytes)
            .ok_or_else(|| unsupported("PowerPoint gradient model budget exceeded"))?;
        let mut stops = Vec::new();
        stops
            .try_reserve_exact(count)
            .map_err(|_| unsupported("PowerPoint gradient model allocation failed"))?;
        for (stop, alpha) in self.projection.stops.into_iter().zip(self.alphas) {
            let mut color = String::new();
            color
                .try_reserve_exact(8)
                .map_err(|_| unsupported("PowerPoint gradient model allocation failed"))?;
            write_color(&mut color, stop.color)?;
            if alpha != 65_536 {
                // ooxml-common GradStop: RRGGBBAA when an alpha applies.
                let byte = ((u64::from(alpha) * 255 + 32_768) / 65_536) as u8;
                write!(&mut color, "{byte:02X}")
                    .map_err(|_| unsupported("PowerPoint gradient color formatting failed"))?;
            }
            stops.push(GradStop {
                position: stop.position(),
                color,
            });
        }
        Ok(Fill::Gradient {
            stops,
            angle: self.projection.angle.degrees(),
            grad_type: "linear".to_owned(),
            scaled: Some(self.scaled),
            path: None,
            fill_to_rect: None,
            tile_rect: None,
            flip: None,
            rot_with_shape: Some(self.rotate_with_shape),
        })
    }
}

fn write_color(output: &mut String, color: u32) -> Result<(), String> {
    write!(
        output,
        "{:02X}{:02X}{:02X}",
        color & 255,
        (color >> 8) & 255,
        (color >> 16) & 255
    )
    .map_err(|_| unsupported("PowerPoint gradient color formatting failed"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn source(stops: &[(u32, u32)]) -> office_gradient::Borrowed<'static> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&(stops.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&(stops.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&8u16.to_le_bytes());
        for &(color, position) in stops {
            bytes.extend_from_slice(&color.to_le_bytes());
            bytes.extend_from_slice(&position.to_le_bytes());
        }
        let bytes = Box::leak(bytes.into_boxed_slice());
        let mut source = office_gradient::Borrowed::default();
        source.set(bytes);
        source
    }

    fn paint() -> Paint {
        Paint {
            fill_type: Some(4),
            fill: Some(0x0033_2211),
            fill_back: Some(0x0066_5544),
            fill_focus: Some(100),
            fill_angle: Some(270 << 16),
            fill_shade_type: Some(0x4000_0003),
            ..Paint::default()
        }
    }

    fn resolved(paint: &Paint, source: &office_gradient::Borrowed<'_>) -> Option<ResolvedGradient> {
        let mut work = usize::MAX;
        let mut bytes = usize::MAX;
        paint
            .project_gradient(source, true, None, &mut work, &mut bytes)
            .unwrap()
    }

    fn attempt(
        paint: &Paint,
        source: &office_gradient::Borrowed<'_>,
    ) -> Result<Option<ResolvedGradient>, String> {
        let (mut work, mut bytes) = (usize::MAX, usize::MAX);
        paint.project_gradient(source, true, None, &mut work, &mut bytes)
    }

    #[test]
    fn admits_linear_shades_and_rejects_unsupported_shade_options() {
        let source = source(&[(0x0033_2211, 0)]);
        assert!(resolved(&paint(), &source).is_some());
        let mut compatible = paint();
        compatible.fill_shade_type = Some(0xffff_ffe3);
        compatible.fill_dztype = Some(7);
        compatible.fill_origins = [Some(1), Some(2), Some(3), Some(4)];
        assert!(resolved(&compatible, &source).is_some());
        for shade in [0, 0x4000_0003] {
            let mut candidate = paint();
            candidate.fill_shade_type = Some(shade);
            candidate.fill_type = Some(7);
            let gradient = resolved(&candidate, &source).unwrap();
            assert!(gradient.scaled);
        }
        // Not a (visible) shade: no gradient, the solid/none path decides.
        for none in [
            |p: &mut Paint| p.fill_type = Some(3),
            |p: &mut Paint| p.filled = Some(false),
            |p: &mut Paint| p.fill_ok = Some(false),
        ] {
            let mut candidate = paint();
            none(&mut candidate);
            assert!(resolved(&candidate, &source).is_none());
        }
        // Shades without evidence fail closed instead of degrading to solid.
        for reject in [
            |p: &mut Paint| p.fill_type = Some(5),
            |p: &mut Paint| p.fill_type = Some(6),
            |p: &mut Paint| p.fill_type = Some(8),
            |p: &mut Paint| p.fill_shape = Some(false),
            |p: &mut Paint| p.fill_rect = Some(true),
            |p: &mut Paint| p.fill_alpha = Some(1),
            |p: &mut Paint| p.fill_back_alpha = Some(1),
            |p: &mut Paint| p.fill_shade_type = Some(2),
        ] {
            let mut candidate = paint();
            reject(&mut candidate);
            assert!(attempt(&candidate, &source).is_err());
        }
        let (mut work, mut bytes) = (usize::MAX, usize::MAX);
        assert!(paint()
            .project_gradient(&source, false, None, &mut work, &mut bytes)
            .unwrap()
            .is_none());
    }

    #[test]
    fn two_colour_shades_keep_stop_origin_opacity_even_for_equal_colours() {
        let mut p = paint();
        p.fill_back = p.fill;
        p.fill_alpha = Some(27_525);
        p.fill_back_alpha = Some(0);
        p.fill_focus = Some(50);
        let Fill::Gradient { stops, scaled, .. } = resolved(&p, &source(&[]))
            .unwrap()
            .into_model(&mut usize::MAX.clone())
            .unwrap()
        else {
            panic!("expected gradient")
        };
        let colors: Vec<_> = stops.iter().map(|stop| stop.color.as_str()).collect();
        // Focus 50: fill, back, fill. 27525/65536 -> 0x6B, 0 -> 0x00.
        assert_eq!(colors, ["1122336B", "11223300", "1122336B"]);
        assert_eq!(scaled, Some(false));
    }

    #[test]
    fn empty_missing_first_and_unknown_scheme_colors() {
        // An empty shade array is a two-colour shade (fill -> back colour).
        let two = resolved(&paint(), &source(&[])).unwrap();
        assert_eq!(two.alphas, [65_536, 65_536]);
        assert!(attempt(&paint(), &source(&[(1, 1)])).is_err());
        assert!(attempt(&paint(), &source(&[(0x0800_0007, 0)])).is_err());
        let scheme = [0x0011_2233; 8];
        let indexed = source(&[(0x0800_0001, 0)]);
        let (mut work, mut bytes) = (usize::MAX, usize::MAX);
        assert!(paint()
            .project_gradient(&indexed, true, Some(&scheme), &mut work, &mut bytes)
            .unwrap()
            .is_some());
    }

    #[test]
    fn explicit_zero_booleans_and_inheritance_are_preserved() {
        let source = source(&[(0x0033_2211, 0)]);
        let mut parent = paint();
        parent.rotate_fill_with_shape = Some(true);
        let local = Paint {
            fill_type: Some(4),
            fill: parent.fill,
            fill_back: parent.fill_back,
            fill_focus: parent.fill_focus,
            fill_angle: parent.fill_angle,
            fill_shade_type: parent.fill_shade_type,
            rotate_fill_with_shape: Some(false),
            ..Paint::default()
        };
        let (mut work, mut bytes) = (usize::MAX, usize::MAX);
        let mut model_bytes = usize::MAX;
        let model = local
            .inherit(&parent)
            .project_gradient(&source, true, None, &mut work, &mut bytes)
            .unwrap()
            .unwrap()
            .into_model(&mut model_bytes)
            .unwrap();
        assert!(matches!(
            model,
            Fill::Gradient {
                rot_with_shape: Some(false),
                ..
            }
        ));
    }

    #[test]
    fn model_keeps_exact_projection_and_rounding_boundaries() {
        let source = source(&[
            (0x0033_2211, 0),
            (0x0066_5544, 32_768),
            (0x0001_0203, 65_536),
        ]);
        let mut model_bytes = usize::MAX;
        let model = resolved(&paint(), &source)
            .unwrap()
            .into_model(&mut model_bytes)
            .unwrap();
        match model {
            Fill::Gradient {
                stops,
                angle,
                grad_type,
                scaled,
                path,
                fill_to_rect,
                tile_rect,
                flip,
                rot_with_shape,
            } => {
                assert_eq!(stops[0].color, "112233");
                assert_eq!(stops[1].position, 0.5);
                assert_eq!(stops.last().unwrap().color, "030201");
                assert_eq!(angle, 180.0);
                assert_eq!(grad_type, "linear");
                assert_eq!(
                    (scaled, path, fill_to_rect, tile_rect, flip),
                    (Some(false), None, None, None, None)
                );
                assert_eq!(rot_with_shape, Some(false));
            }
            _ => panic!("expected gradient"),
        }
    }

    #[test]
    fn decoder_projection_and_adapter_budgets_fail_closed() {
        let source = source(&[(0x0033_2211, 0)]);
        for focus in [i32::MIN, i32::MAX] {
            let mut invalid = paint();
            invalid.fill_focus = Some(focus as u32);
            let (mut work, mut bytes) = (usize::MAX, usize::MAX);
            assert!(invalid
                .project_gradient(&source, true, None, &mut work, &mut bytes)
                .is_err());
        }
        let mut bytes = usize::MAX;
        assert!(paint()
            .project_gradient(&source, true, None, &mut 0, &mut bytes)
            .is_err());
        let mut work = usize::MAX;
        assert!(paint()
            .project_gradient(&source, true, None, &mut work, &mut 0)
            .is_err());
        // The shared budget is not split into arbitrary scratch/output halves:
        // output slots are larger, so an exact combined budget must succeed.
        let (mut work, mut bytes) = (usize::MAX, usize::MAX);
        paint()
            .project_gradient(&source, true, None, &mut work, &mut bytes)
            .unwrap()
            .unwrap();
        let projection_bytes = usize::MAX - bytes;
        let mut exact_bytes = projection_bytes;
        let mut work = usize::MAX;
        assert!(paint()
            .project_gradient(&source, true, None, &mut work, &mut exact_bytes)
            .unwrap()
            .is_some());
        assert_eq!(exact_bytes, 0);
        assert!(paint()
            .project_gradient(&source, true, None, &mut work, &mut (projection_bytes - 1))
            .is_err());
        let descriptor = resolved(&paint(), &source).unwrap();
        let mut unlimited = usize::MAX;
        descriptor.into_model(&mut unlimited).unwrap();
        let model_len = usize::MAX - unlimited;
        assert!(resolved(&paint(), &source)
            .unwrap()
            .into_model(&mut (model_len - 1))
            .is_err());
        let mut exact = model_len;
        assert!(resolved(&paint(), &source)
            .unwrap()
            .into_model(&mut exact)
            .is_ok());
        assert_eq!(exact, 0);
    }
}
