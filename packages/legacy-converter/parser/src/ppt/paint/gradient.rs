//! Adapter from retained OfficeArt shade facts to DrawingML/native gradients.

use super::Paint;
use crate::officeart::gradient as office_gradient;
use crate::ppt::{scheme, unsupported};
use pptx_model::{Fill, GradStop};
use std::fmt::Write;
use std::mem::size_of;

pub(in crate::ppt) struct ResolvedGradient {
    projection: office_gradient::projection::Projection,
    rotate_with_shape: bool,
}

impl Paint {
    pub(in crate::ppt) fn project_gradient(
        &self,
        source: &office_gradient::Borrowed<'_>,
        allow_fill: bool,
        colors: Option<&scheme::Scheme>,
        work_budget: &mut usize,
        byte_budget: &mut usize,
    ) -> Result<Option<ResolvedGradient>, String> {
        if self.fill_type != Some(4)
            || !allow_fill
            || !self.filled.unwrap_or(true)
            || !self.fill_ok.unwrap_or(true)
            || !self.fill_shape.unwrap_or(true)
            || self.fill_rect.unwrap_or(false)
            || self.fill_alpha.unwrap_or(65_536) != 65_536
            || self.fill_back_alpha.unwrap_or(65_536) != 65_536
            || self.fill_shade_type.unwrap_or(0x4000_0003) & 0x1f != 3
        {
            return Ok(None);
        }

        // The decoded vector and projection scratch coexist. Reserve the exact
        // normalized and projected vector payloads before either allocation.
        let Some(authored) = source.decode(work_budget, byte_budget)? else {
            return Ok(None);
        };
        if authored.is_empty() || authored[0].position != 0 {
            return Ok(None);
        }
        let focus = self.fill_focus.unwrap_or(0) as i32;
        let requirements = office_gradient::projection::requirements(&authored, focus)?;
        let mut projection_scratch = requirements.scratch_bytes;
        let mut projection_output = requirements.output_bytes;
        let projection_bytes = projection_scratch
            .checked_add(projection_output)
            .ok_or_else(|| unsupported("PowerPoint gradient byte budget overflow"))?;
        *byte_budget = byte_budget
            .checked_sub(projection_bytes)
            .ok_or_else(|| unsupported("PowerPoint gradient byte budget exceeded"))?;
        let mut projection = office_gradient::projection::project(
            &authored,
            self.fill.unwrap_or(0x00ff_ffff),
            self.fill_back.unwrap_or(0x00ff_ffff),
            focus,
            self.fill_angle.unwrap_or(0) as i32,
            work_budget,
            &mut projection_scratch,
            &mut projection_output,
        )?;
        debug_assert_eq!((projection_scratch, projection_output), (0, 0));

        for stop in &mut projection.stops {
            let Some(color) = scheme::drawing(stop.color, colors) else {
                return Ok(None);
            };
            stop.color = color;
        }
        Ok(Some(ResolvedGradient {
            projection,
            rotate_with_shape: self.rotate_fill_with_shape.unwrap_or(false),
        }))
    }
}

impl ResolvedGradient {
    pub(in crate::ppt) fn to_model(self, byte_budget: &mut usize) -> Result<Fill, String> {
        let count = self.projection.stops.len();
        let bytes = count
            .checked_mul(size_of::<GradStop>() + 6)
            .and_then(|value| value.checked_add("linear".len()))
            .ok_or_else(|| unsupported("PowerPoint gradient model budget overflow"))?;
        *byte_budget = byte_budget
            .checked_sub(bytes)
            .ok_or_else(|| unsupported("PowerPoint gradient model budget exceeded"))?;
        let mut stops = Vec::new();
        stops
            .try_reserve_exact(count)
            .map_err(|_| unsupported("PowerPoint gradient model allocation failed"))?;
        for stop in self.projection.stops {
            let mut color = String::new();
            color
                .try_reserve_exact(6)
                .map_err(|_| unsupported("PowerPoint gradient model allocation failed"))?;
            write_color(&mut color, stop.color)?;
            stops.push(GradStop {
                position: stop.position(),
                color,
            });
        }
        Ok(Fill::Gradient {
            stops,
            angle: self.projection.angle.degrees(),
            grad_type: "linear".to_owned(),
            scaled: None,
            path: None,
            fill_to_rect: None,
            tile_rect: None,
            flip: None,
            rot_with_shape: Some(self.rotate_with_shape),
        })
    }

    pub(in crate::ppt) fn to_xml(&self, byte_budget: &mut usize) -> Result<String, String> {
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

    #[test]
    fn admits_only_the_bounded_linear_opaque_shape_fill() {
        let source = source(&[(0x0033_2211, 0)]);
        assert!(resolved(&paint(), &source).is_some());
        let mut compatible = paint();
        compatible.fill_shade_type = Some(0xffff_ffe3);
        compatible.fill_dztype = Some(7);
        compatible.fill_origins = [Some(1), Some(2), Some(3), Some(4)];
        assert!(resolved(&compatible, &source).is_some());
        for reject in [
            |p: &mut Paint| p.fill_type = Some(3),
            |p: &mut Paint| p.filled = Some(false),
            |p: &mut Paint| p.fill_ok = Some(false),
            |p: &mut Paint| p.fill_shape = Some(false),
            |p: &mut Paint| p.fill_rect = Some(true),
            |p: &mut Paint| p.fill_alpha = Some(1),
            |p: &mut Paint| p.fill_back_alpha = Some(1),
            |p: &mut Paint| p.fill_shade_type = Some(2),
        ] {
            let mut candidate = paint();
            reject(&mut candidate);
            assert!(resolved(&candidate, &source).is_none());
        }
        let (mut work, mut bytes) = (usize::MAX, usize::MAX);
        assert!(paint()
            .project_gradient(&source, false, None, &mut work, &mut bytes)
            .unwrap()
            .is_none());
    }

    #[test]
    fn empty_missing_first_and_unknown_scheme_colors_are_not_admitted() {
        assert!(resolved(&paint(), &source(&[])).is_none());
        assert!(resolved(&paint(), &source(&[(1, 1)])).is_none());
        assert!(resolved(&paint(), &source(&[(0x0800_0007, 0)])).is_none());
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
                    (None, None, None, None, None)
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
