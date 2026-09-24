//! Adapter from retained OfficeArt shade facts to DrawingML/native gradients.

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

// Two-colour shades are projected with these marker colours so each output
// stop keeps its origin (fill or back colour) even when both colours match.
const FRONT_MARKER: u32 = 0;
const BACK_MARKER: u32 = 1;

impl Paint {
    /// Linear OfficeArt shades -> DrawingML `a:lin` gradients.
    ///
    /// Evidence for the mapping beyond MS-ODRAW 2.4.13 (whose shade
    /// illustrations do not state DrawingML equivalents): every shape-level
    /// gradient in the local PowerPoint corpus that also carries a metroBlob
    /// (MS-ODRAW 2.3.4.41, PowerPoint's own DrawingML for the same shape) was
    /// paired with its binary fill properties:
    /// - msofillShade (4) is `a:lin scaled="0"` and msofillShadeScale (7) is
    ///   `a:lin scaled="1"`; both use `ang` = 90 degrees - fillAngle.
    /// - fillShadeType 0 and the default 0x40000003 (gamma + sigma) produce the
    ///   same DrawingML stops; other shade types stay unsupported.
    /// - Without fillShadeColors, fillOpacity applies to every stop taken from
    ///   the fill colour and fillBackOpacity to every stop taken from the back
    ///   colour (e.g. 27525/65536 -> alpha 42000, back 0 -> alpha 0).
    /// Opacity combined with an authored shade-colour array has no evidence
    /// and is rejected, as are the path shades (5, 6) and the host-defined
    /// title shade (8).
    pub(in crate::ppt) fn project_gradient(
        &self,
        source: &office_gradient::Borrowed<'_>,
        allow_fill: bool,
        colors: Option<&scheme::Scheme>,
        work_budget: &mut usize,
        byte_budget: &mut usize,
    ) -> Result<Option<ResolvedGradient>, String> {
        let scaled = match self.fill_type {
            Some(4) => false,
            Some(7) => true,
            Some(5 | 6 | 8) => {
                return Err(unsupported(
                    "PowerPoint path or title gradient fills are not supported yet",
                ))
            }
            _ => return Ok(None),
        };
        if !allow_fill || !self.filled.unwrap_or(true) || !self.fill_ok.unwrap_or(true) {
            return Ok(None);
        }
        if !self.fill_shape.unwrap_or(true)
            || self.fill_rect.unwrap_or(false)
            || !matches!(self.fill_shade_type.unwrap_or(0x4000_0003) & 0x1f, 0 | 3)
        {
            return Err(unsupported(
                "PowerPoint gradient shade options are not supported yet",
            ));
        }
        let front_alpha = self.fill_alpha.unwrap_or(65_536);
        let back_alpha = self.fill_back_alpha.unwrap_or(65_536);

        // The decoded vector and projection scratch coexist. Reserve the exact
        // normalized and projected vector payloads before either allocation.
        let authored = source.decode(work_budget, byte_budget)?.unwrap_or_default();
        if !authored.is_empty() && (front_alpha != 65_536 || back_alpha != 65_536) {
            return Err(unsupported(
                "PowerPoint gradient opacity with shade colours is not supported yet",
            ));
        }
        if authored.first().is_some_and(|stop| stop.position != 0) {
            return Err(unsupported(
                "PowerPoint gradient shade colours must start at position 0",
            ));
        }
        let two_colour = authored.is_empty();
        let focus = self.fill_focus.unwrap_or(0) as i32;
        let requirements = office_gradient::projection::requirements(&authored, focus)?;
        let mut projection_scratch = requirements.scratch_bytes;
        let mut projection_output = requirements.output_bytes;
        let projection_bytes = projection_scratch
            .checked_add(projection_output)
            .and_then(|bytes| bytes.checked_add(requirements.output_bytes))
            .ok_or_else(|| unsupported("PowerPoint gradient byte budget overflow"))?;
        *byte_budget = byte_budget
            .checked_sub(projection_bytes)
            .ok_or_else(|| unsupported("PowerPoint gradient byte budget exceeded"))?;
        let (front, back) = if two_colour {
            (FRONT_MARKER, BACK_MARKER)
        } else {
            (
                self.fill.unwrap_or(0x00ff_ffff),
                self.fill_back.unwrap_or(0x00ff_ffff),
            )
        };
        let mut projection = office_gradient::projection::project(
            &authored,
            front,
            back,
            focus,
            self.fill_angle.unwrap_or(0) as i32,
            work_budget,
            &mut projection_scratch,
            &mut projection_output,
        )?;
        debug_assert_eq!((projection_scratch, projection_output), (0, 0));

        let mut alphas = Vec::new();
        alphas
            .try_reserve_exact(projection.stops.len())
            .map_err(|_| unsupported("PowerPoint gradient allocation failed"))?;
        for stop in &mut projection.stops {
            let (color, alpha) = match (two_colour, stop.color) {
                (true, FRONT_MARKER) => (self.fill.unwrap_or(0x00ff_ffff), front_alpha),
                (true, _) => (self.fill_back.unwrap_or(0x00ff_ffff), back_alpha),
                (false, color) => (color, 65_536),
            };
            let Some(color) = scheme::drawing(color, colors) else {
                return Err(unsupported("PowerPoint gradient colour cannot be resolved"));
            };
            stop.color = color;
            alphas.push(alpha);
        }
        Ok(Some(ResolvedGradient {
            projection,
            alphas,
            scaled,
            rotate_with_shape: self.rotate_fill_with_shape.unwrap_or(false),
        }))
    }
}

impl ResolvedGradient {
    pub(in crate::ppt) fn to_model(self, byte_budget: &mut usize) -> Result<Fill, String> {
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

    pub(in crate::ppt) fn to_xml(&self, byte_budget: &mut usize) -> Result<String, String> {
        // The withdrawn OOXML route only ever emitted opaque, unscaled shades.
        if self.scaled || self.alphas.iter().any(|alpha| *alpha != 65_536) {
            return Err(unsupported(
                "PowerPoint scaled or translucent gradient has no OOXML route",
            ));
        }
        const OPEN: &str = "<a:gradFill rotWithShape=\"";
        const LIST: &str = "\"><a:gsLst>";
        const STOP_OPEN: &str = "<a:gs pos=\"";
        const STOP_COLOR: &str = "\"><a:srgbClr val=\"";
        const STOP_CLOSE: &str = "\"/></a:gs>";
        const CLOSE: &str = "</a:gsLst><a:lin ang=\"";
        const END: &str = "\"/></a:gradFill>";
        let mut length = OPEN.len() + 1 + LIST.len() + CLOSE.len() + END.len();
        length = length
            .checked_add(decimal_len(self.projection.angle.angle_units() as u32))
            .ok_or_else(|| unsupported("PowerPoint gradient XML budget overflow"))?;
        for stop in &self.projection.stops {
            length = length
                .checked_add(
                    STOP_OPEN.len()
                        + decimal_len(stop.position_units())
                        + STOP_COLOR.len()
                        + 6
                        + STOP_CLOSE.len(),
                )
                .ok_or_else(|| unsupported("PowerPoint gradient XML budget overflow"))?;
        }
        *byte_budget = byte_budget
            .checked_sub(length)
            .ok_or_else(|| unsupported("PowerPoint gradient XML budget exceeded"))?;
        let mut output = String::new();
        output
            .try_reserve_exact(length)
            .map_err(|_| unsupported("PowerPoint gradient XML allocation failed"))?;
        output.push_str(OPEN);
        output.push(if self.rotate_with_shape { '1' } else { '0' });
        output.push_str(LIST);
        for stop in &self.projection.stops {
            output.push_str(STOP_OPEN);
            write!(&mut output, "{}", stop.position_units())
                .map_err(|_| unsupported("PowerPoint gradient XML formatting failed"))?;
            output.push_str(STOP_COLOR);
            write_color(&mut output, stop.color)?;
            output.push_str(STOP_CLOSE);
        }
        output.push_str(CLOSE);
        write!(&mut output, "{}", self.projection.angle.angle_units())
            .map_err(|_| unsupported("PowerPoint gradient XML formatting failed"))?;
        output.push_str(END);
        debug_assert_eq!(output.len(), length);
        Ok(output)
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

fn decimal_len(mut value: u32) -> usize {
    let mut result = 1;
    while value >= 10 {
        result += 1;
        value /= 10;
    }
    result
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
        let mut result = Paint::default();
        result.fill_type = Some(4);
        result.fill = Some(0x0033_2211);
        result.fill_back = Some(0x0066_5544);
        result.fill_focus = Some(100);
        result.fill_angle = Some(270 << 16);
        result.fill_shade_type = Some(0x4000_0003);
        result
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
            .to_model(&mut usize::MAX.clone())
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
        let mut local = Paint::default();
        local.fill_type = Some(4);
        local.fill = parent.fill;
        local.fill_back = parent.fill_back;
        local.fill_focus = parent.fill_focus;
        local.fill_angle = parent.fill_angle;
        local.fill_shade_type = parent.fill_shade_type;
        local.rotate_fill_with_shape = Some(false);
        let (mut work, mut bytes) = (usize::MAX, usize::MAX);
        let mut model_bytes = usize::MAX;
        let model = local
            .inherit(&parent)
            .project_gradient(&source, true, None, &mut work, &mut bytes)
            .unwrap()
            .unwrap()
            .to_model(&mut model_bytes)
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
    fn model_and_xml_share_exact_projection_and_rounding_boundaries() {
        let source = source(&[
            (0x0033_2211, 0),
            (0x0066_5544, 32_768),
            (0x0001_0203, 65_536),
        ]);
        let descriptor = resolved(&paint(), &source).unwrap();
        let mut xml_bytes = usize::MAX;
        let xml = descriptor.to_xml(&mut xml_bytes).unwrap();
        assert!(xml.contains("pos=\"50000\""));
        assert!(xml.contains("ang=\"10800000\""));
        assert!(xml.contains("val=\"112233\""));
        let mut model_bytes = usize::MAX;
        let model = resolved(&paint(), &source)
            .unwrap()
            .to_model(&mut model_bytes)
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
        let xml_len = descriptor.to_xml(&mut unlimited).unwrap().len();
        assert!(descriptor.to_xml(&mut (xml_len - 1)).is_err());
        assert_eq!(
            descriptor.to_xml(&mut xml_len.clone()).unwrap().len(),
            xml_len
        );
        let descriptor = resolved(&paint(), &source).unwrap();
        let mut unlimited = usize::MAX;
        descriptor.to_model(&mut unlimited).unwrap();
        let model_len = usize::MAX - unlimited;
        assert!(resolved(&paint(), &source)
            .unwrap()
            .to_model(&mut (model_len - 1))
            .is_err());
        let mut exact = model_len;
        assert!(resolved(&paint(), &source)
            .unwrap()
            .to_model(&mut exact)
            .is_ok());
        assert_eq!(exact, 0);
    }
}
