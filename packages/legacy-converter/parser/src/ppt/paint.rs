//! OfficeArt paint projected into presentation-model fills and strokes,
//! without renderer extensions.
use super::scheme;
use pptx_model::{ArrowEnd, Fill, Stroke};
mod gradient;

pub(super) use crate::officeart::paint::Paint;

impl Paint {
    /// Plain foreground msofillPattern (MS-ODRAW 2.4.11): the fillBlip BLIP
    /// is the pattern, fillColor its foreground and fillBackColor its
    /// background (2.3.7.2 table). Returns the BLIP, both colours and their
    /// opacities, under the same placement vetoes as a picture fill.
    pub(super) fn pattern_image(&self) -> Option<(u32, u32, u32, u32, u32)> {
        (self.fill_type == Some(1)
            && self.fill_blip.unwrap_or(0) != 0
            && !self.fill_rect.unwrap_or(false)
            && self.fill_shape.unwrap_or(true)
            && !self.rotate_fill_with_shape.unwrap_or(false)
            && self.fill_dztype.unwrap_or(0) == 0
            && self
                .fill_origins
                .iter()
                .all(|value| value.unwrap_or(0) == 0)
            && self.filled.unwrap_or(true)
            && self.fill_ok.unwrap_or(true))
        .then_some((
            self.fill_blip.unwrap_or(0),
            self.fill.unwrap_or(0xffffff),
            self.fill_alpha.unwrap_or(65536),
            self.fill_back.unwrap_or(0xffffff),
            self.fill_back_alpha.unwrap_or(65536),
        ))
    }

    /// Picture-frame backing fill. PowerPoint 16 writes the ~2,000 corpus
    /// picture frames that do not set fFilled themselves as `noFill` (even
    /// though the drawing-group defaults set fFilled), and the frames with an
    /// explicit fFilled = 1 and fillColor as a solid spPr fill, which its PDF
    /// export paints behind transparent pixels. Only that evidenced case is
    /// projected: `Ok(None)` for an unfilled frame, `Err` for a filled frame
    /// whose fill type or colour source has no evidence.
    pub(super) fn picture_backing(&self) -> Result<Option<(u32, u32)>, String> {
        if self.filled != Some(true) || !self.fill_ok.unwrap_or(true) {
            return Ok(None);
        }
        match (self.fill_type.unwrap_or(0), self.fill) {
            (0, Some(color)) => Ok(Some((color, self.fill_alpha.unwrap_or(65536)))),
            (0, None) => Err(super::unsupported(
                "PowerPoint picture frame fill without a color",
            )),
            _ => Err(super::unsupported(
                "PowerPoint picture frame non-solid fill",
            )),
        }
    }

    pub(super) fn model_with_custom_geometry(
        &self,
        scheme: Option<&scheme::Scheme>,
        allow_fill: bool,
        allow_line: bool,
        image_fill: Option<Fill>,
    ) -> (Option<Fill>, Option<Stroke>) {
        let image_fill = image_fill.filter(|_| self.foreground_image().is_some());
        let fill = if allow_fill {
            image_fill.or_else(|| {
                self.solid_fill_values(allow_fill)
                    .and_then(|(color, alpha)| model_solid(color, alpha, scheme))
            })
        } else {
            None
        };
        let stroke = self
            .solid_line_values(allow_line)
            .and_then(|_| self.model_stroke(scheme));
        (Some(fill.unwrap_or(Fill::None)), stroke)
    }

    pub(super) fn background_model(
        &self,
        scheme: Option<&scheme::Scheme>,
        image_fill: Option<Fill>,
    ) -> Option<Fill> {
        if !self.filled.unwrap_or(true) || !self.fill_ok.unwrap_or(true) {
            return Some(Fill::None);
        }
        if self.fill_rect.unwrap_or(false) {
            return None;
        }
        if self.fill_type.unwrap_or(0) == 3 {
            return image_fill.filter(|_| self.background_image().is_some());
        }
        (self.fill_type.unwrap_or(0) == 0)
            .then(|| {
                model_solid(
                    self.fill.unwrap_or(0xffffff),
                    self.fill_alpha.unwrap_or(65536),
                    scheme,
                )
            })
            .flatten()
    }

    fn model_stroke(&self, scheme: Option<&scheme::Scheme>) -> Option<Stroke> {
        let color = model_color(
            self.line.unwrap_or(0),
            self.line_alpha.unwrap_or(65536),
            scheme,
        )?;
        let (line_join, miter_limit) = self.details.join();
        let arrow = |end: crate::officeart::stroke::LineEnd<'_>| ArrowEnd {
            kind: end.kind.to_owned(),
            w: end.width.to_owned(),
            len: end.length.to_owned(),
        };
        Some(Stroke {
            color,
            width: i64::from(self.width.unwrap_or(9525)),
            fill: None,
            dash_style: self
                .dash
                .and_then(crate::officeart::stroke::preset_dash)
                .filter(|value| *value != "solid")
                .map(str::to_owned),
            custom_dash: Vec::new(),
            line_cap: Some(self.details.canvas_cap().to_owned()),
            line_join: Some(line_join.to_owned()),
            miter_limit,
            alignment: None,
            head_end: self.details.line_end(0).map(arrow),
            tail_end: self.details.line_end(1).map(arrow),
            cmpd: None,
        })
    }
}

pub(super) fn model_solid(
    color: u32,
    opacity: u32,
    scheme: Option<&scheme::Scheme>,
) -> Option<Fill> {
    model_color(color, opacity, scheme).map(|color| Fill::Solid { color })
}

pub(super) fn model_color(
    color: u32,
    opacity: u32,
    scheme: Option<&scheme::Scheme>,
) -> Option<String> {
    let color = scheme::drawing(color, scheme)?;
    let mut result = format!(
        "{:02X}{:02X}{:02X}",
        color & 255,
        (color >> 8) & 255,
        (color >> 16) & 255
    );
    if opacity != 65536 {
        let alpha = (u64::from(opacity) * 255 + 32768) / 65536;
        result.push_str(&format!("{alpha:02X}"));
    }
    Some(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The direct model's paint for a filled preset (allow_fill) or a line
    /// preset (no fill area); lines are always allowed.
    fn model(p: &Paint, allow_fill: bool) -> (Option<Fill>, Option<Stroke>) {
        p.model_with_custom_geometry(None, allow_fill, true, None)
    }

    fn solid(fill: &Option<Fill>) -> Option<&str> {
        match fill {
            Some(Fill::Solid { color }) => Some(color),
            _ => None,
        }
    }

    #[test]
    fn explicit_paint_vetoes_suppress_solid_fill_and_line() {
        for allow in [false, true] {
            for enabled in [None, Some(false), Some(true)] {
                for ok in [None, Some(false), Some(true)] {
                    for kind in [None, Some(0), Some(3), Some(7)] {
                        let p = Paint {
                            fill: Some(0x332211),
                            line: Some(0x665544),
                            filled: enabled,
                            lined: enabled,
                            fill_ok: ok,
                            line_ok: ok,
                            fill_type: kind,
                            line_type: kind,
                            ..Paint::default()
                        };
                        let painted = allow
                            && enabled != Some(false)
                            && ok != Some(false)
                            && kind.unwrap_or(0) == 0;
                        let (fill, line) = p.model_with_custom_geometry(None, allow, allow, None);
                        assert_eq!(solid(&fill), painted.then_some("112233"));
                        assert_eq!(
                            line.map(|line| line.color),
                            painted.then(|| "445566".into())
                        );
                    }
                }
            }
        }
    }

    #[test]
    fn direct_model_preserves_solid_opacity_and_stroke_semantics() {
        let mut p = Paint::default();
        for (id, value) in [
            (0x181, 0x332211),
            (0x182, 32768),
            (0x1c0, 0x665544),
            (0x1c1, 16384),
            (0x1cb, 25400),
            (0x1ce, 8),
            (0x1d0, 1),
            (0x1d1, 5),
            (0x1d6, 1),
            (0x1cc, 0x18000),
            (0x1d7, 1),
        ] {
            p.property(id, value).unwrap();
        }
        let (fill, stroke) = model(&p, true);
        assert_eq!(solid(&fill), Some("11223380"));
        let stroke = stroke.unwrap();
        assert_eq!(stroke.color, "44556640");
        assert_eq!(stroke.width, 25400);
        assert_eq!(stroke.dash_style.as_deref(), Some("dashDot"));
        assert_eq!(stroke.line_cap.as_deref(), Some("square"));
        assert_eq!(stroke.line_join.as_deref(), Some("miter"));
        assert_eq!(stroke.miter_limit, Some(1.5));
        assert_eq!(stroke.head_end.unwrap().kind, "triangle");
        assert_eq!(stroke.tail_end.unwrap().kind, "arrow");
    }

    #[test]
    fn direct_model_keeps_no_fill_and_unsupported_paint_absent() {
        let mut p = Paint::default();
        p.property(0x181, 0x332211).unwrap();
        p.property(0x1c0, 0x665544).unwrap();
        p.property(0x1bf, 0x00100000).unwrap();
        p.property(0x1c4, 1).unwrap();
        let (fill, stroke) = model(&p, true);
        assert!(matches!(fill, Some(Fill::None)));
        assert!(stroke.is_none());

        let mut unresolved = Paint::default();
        unresolved.property(0x181, 0x08000001).unwrap();
        let (fill, _) = model(&unresolved, true);
        assert!(matches!(fill, Some(Fill::None)));
        assert!(unresolved.background_model(None, None).is_none());
    }

    #[test]
    fn direct_background_and_passive_image_use_only_provided_model_fill() {
        let mut p = Paint::default();
        p.property(0x180, 3).unwrap();
        p.property(0x4186, 9).unwrap();
        assert!(p.background_model(None, None).is_none());
        assert!(matches!(
            p.background_model(None, Some(Fill::None)),
            Some(Fill::None)
        ));
        let (fill, _) = p.model_with_custom_geometry(None, true, false, Some(Fill::None));
        assert!(matches!(fill, Some(Fill::None)));

        p.property(0x1bf, 0x00100000).unwrap();
        assert!(p
            .background_model(None, Some(Fill::None))
            .is_some_and(|fill| matches!(fill, Fill::None)));
        let (fill, _) = p.model_with_custom_geometry(
            None,
            true,
            false,
            Some(Fill::Solid {
                color: "BADBAD".into(),
            }),
        );
        assert!(matches!(fill, Some(Fill::None)));
    }

    #[test]
    fn all_dash_presets_retain_the_line_and_inherit_explicit_solid() {
        for (value, name) in [
            "solid",
            "sysDash",
            "sysDot",
            "sysDashDot",
            "sysDashDotDot",
            "dot",
            "dash",
            "lgDash",
            "dashDot",
            "lgDashDot",
            "lgDashDotDot",
        ]
        .iter()
        .enumerate()
        {
            let mut parent = Paint::default();
            parent.property(0x1c0, 0xff0000).unwrap();
            parent.property(0x1ce, value as u32).unwrap();
            let inherited = model(&Paint::default().inherit(&parent), false).1.unwrap();
            assert_eq!(inherited.color, "0000FF");
            // A solid dash is the model's default line, not a named dash.
            assert_eq!(
                inherited.dash_style.as_deref(),
                (value != 0).then_some(*name)
            );
            assert_eq!(inherited.line_join.as_deref(), Some("round"));
            let mut child = Paint::default();
            child.property(0x1ce, 0).unwrap();
            let explicit = model(&child.inherit(&parent), false).1.unwrap();
            assert_eq!(explicit.dash_style, None);
            child.property(0x1ff, 0x00080000).unwrap();
            assert!(model(&child.inherit(&parent), true).1.is_none());
        }
        assert!(Paint::default().property(0x1ce, 11).is_err());
    }

    #[test]
    fn line_decorations_caps_and_joins_reach_the_model() {
        let mut p = Paint::default();
        for (id, value) in [
            (0x1c0, 0),
            (0x1d0, 1),
            (0x1d1, 5),
            (0x1d2, 0),
            (0x1d3, 2),
            (0x1d4, 2),
            (0x1d5, 0),
            (0x1d6, 0),
            (0x1d7, 0),
        ] {
            p.property(id, value).unwrap();
        }
        let stroke = model(&p, false).1.unwrap();
        assert_eq!(stroke.line_cap.as_deref(), Some("round"));
        assert_eq!(stroke.line_join.as_deref(), Some("bevel"));
        let head = stroke.head_end.unwrap();
        assert_eq!(
            (head.kind.as_str(), head.w.as_str(), head.len.as_str()),
            ("triangle", "sm", "lg")
        );
        let tail = stroke.tail_end.unwrap();
        assert_eq!(
            (tail.kind.as_str(), tail.w.as_str(), tail.len.as_str()),
            ("arrow", "lg", "sm")
        );
        p.property(0x1ff, 0x00080000).unwrap();
        assert!(model(&p, true).1.is_none());
    }

    #[test]
    fn decorations_inherit_with_explicit_none_and_miter_defaults() {
        let mut parent = Paint::default();
        for (id, value) in [(0x1c0, 0), (0x1d0, 3), (0x1d1, 4), (0x1d6, 1), (0x1d7, 1)] {
            parent.property(id, value).unwrap();
        }
        let mut child = Paint::default();
        child.property(0x1d0, 0).unwrap();
        child.property(0x1d4, 0).unwrap();
        let stroke = model(&child.inherit(&parent), false).1.unwrap();
        assert!(stroke.head_end.is_none());
        let tail = stroke.tail_end.unwrap();
        assert_eq!(
            (tail.kind.as_str(), tail.w.as_str(), tail.len.as_str()),
            ("oval", "sm", "med")
        );
        assert_eq!(stroke.line_cap.as_deref(), Some("square"));
        assert_eq!(stroke.line_join.as_deref(), Some("miter"));
        assert_eq!(stroke.miter_limit, Some(8.0));
        child.property(0x1cc, 0x00018000).unwrap();
        assert_eq!(
            model(&child.inherit(&parent), false).1.unwrap().miter_limit,
            Some(1.5)
        );
    }

    #[test]
    fn arrow_editability_flags_do_not_hide_end_decorations() {
        let mut p = Paint::default();
        p.property(0x1d1, 1).unwrap();
        // MS-ODRAW 2.3.8.38: fArrowheadsOK controls editing, not rendering.
        p.property(0x1ff, 0x00100000).unwrap();
        let stroke = model(&p, false).1.unwrap();
        assert_eq!(stroke.tail_end.unwrap().kind, "triangle");
        assert_eq!(stroke.line_join.as_deref(), Some("round"));
    }

    #[test]
    fn inherited_paint_keeps_values_independent_from_boolean_use_bits() {
        let mut master = Paint::default();
        for (id, value) in [
            (0x181, 255),
            (0x182, 32768),
            (0x1c0, 0xff0000),
            (0x1cb, 25400),
        ] {
            master.property(id, value).unwrap();
        }
        let mut local = Paint::default();
        local.property(0x1bf, 0).unwrap(); // Unused false is not an override.
        let (fill, stroke) = model(&local.inherit(&master), true);
        assert_eq!(solid(&fill), Some("FF000080"));
        let stroke = stroke.unwrap();
        assert_eq!((stroke.color.as_str(), stroke.width), ("0000FF", 25400));
        local.property(0x1bf, 0x00100000).unwrap();
        local.property(0x1ff, 0x00080000).unwrap();
        let (fill, stroke) = model(&local.inherit(&master), true);
        assert!(matches!(fill, Some(Fill::None)));
        assert!(stroke.is_none());
        local.property(0x1bf, 0x00100010).unwrap();
        local.property(0x181, 0xff00).unwrap();
        let (fill, stroke) = model(&local.inherit(&master), true);
        assert_eq!(solid(&fill), Some("00FF0080"));
        assert!(stroke.is_none());
    }

    #[test]
    fn unsupported_inherited_paint_is_not_replaced_with_solid_defaults() {
        let mut master = Paint::default();
        master.property(0x180, 4).unwrap(); // Gradient.
        master.property(0x1c0, 255).unwrap();
        master.property(0x1ce, 1).unwrap(); // Supported dashed outline, independent of fill.
        let mut local = Paint::default();
        local.property(0x181, 0xff00).unwrap();
        let (fill, stroke) = model(&local.inherit(&master), true);
        assert!(matches!(fill, Some(Fill::None)));
        assert_eq!(stroke.unwrap().dash_style.as_deref(), Some("sysDash"));
        local.property(0x180, 0).unwrap();
        assert_eq!(
            solid(&model(&local.inherit(&master), true).0),
            Some("00FF00")
        );
        master.property(0x1bf, 0x00020002).unwrap();
        assert_eq!(solid(&model(&local.inherit(&master), true).0), None);
        local.property(0x1bf, 0x00020000).unwrap(); // Explicitly clear inherited fill rectangle.
        assert_eq!(
            solid(&model(&local.inherit(&master), true).0),
            Some("00FF00")
        );
    }

    #[test]
    fn gradient_scalars_retain_raw_values_inherit_and_preserve_explicit_zero() {
        let fields = [
            (0x183, 0x0800_0003),
            (0x184, 32768),
            (0x18b, (-90i32 << 16) as u32),
            (0x18c, (-25i32) as u32),
            (0x19c, 0x4000_0003),
        ];
        let mut parent = Paint::default();
        for (id, value) in fields {
            parent.property(id, value).unwrap();
        }
        let inherited = Paint::default().inherit(&parent);
        assert_eq!(inherited.fill_back, Some(0x0800_0003));
        assert_eq!(inherited.fill_back_alpha, Some(32768));
        assert_eq!(inherited.fill_angle, Some((-90i32 << 16) as u32));
        assert_eq!(inherited.fill_focus, Some((-25i32) as u32));
        assert_eq!(inherited.fill_shade_type, Some(0x4000_0003));

        let mut child = Paint::default();
        for (id, _) in fields {
            child.property(id, 0).unwrap();
        }
        let explicit = child.inherit(&parent);
        assert_eq!(explicit.fill_back, Some(0));
        assert_eq!(explicit.fill_back_alpha, Some(0));
        assert_eq!(explicit.fill_angle, Some(0));
        assert_eq!(explicit.fill_focus, Some(0));
        assert_eq!(explicit.fill_shade_type, Some(0));
    }

    #[test]
    fn inherited_scheme_color_resolves_at_destination_not_master() {
        let mut master = Paint::default();
        master.property(0x181, 0x08000004).unwrap();
        let inherited = Paint::default().inherit(&master);
        let mut scheme = [0; 8];
        scheme[4] = 0x563412;
        let fill = |scheme: &scheme::Scheme| {
            inherited
                .model_with_custom_geometry(Some(scheme), true, true, None)
                .0
        };
        assert_eq!(solid(&fill(&scheme)), Some("123456"));
        scheme[4] = 0xabcdef;
        assert_eq!(solid(&fill(&scheme)), Some("EFCDAB"));
    }

    #[test]
    fn background_fill_ignores_line_and_geometry_and_respects_use_bits() {
        let mut p = Paint::default();
        p.property(0x181, 0x112233).unwrap();
        p.property(0x182, 32768).unwrap();
        p.property(0x1c0, 0xff).unwrap();
        p.custom_geometry = true;
        assert!(matches!(
            p.background_model(None, None),
            Some(Fill::Solid { ref color }) if color == "33221180"
        ));
        p.property(0x180, 3).unwrap();
        p.property(0x4186, 9).unwrap();
        assert_eq!(p.background_image(), Some((9, 32768)));
        assert!(p.background_model(None, None).is_none());
        p.property(0x1bf, 0x00100000).unwrap();
        assert!(p.background_image().is_none());
        assert!(matches!(p.background_model(None, None), Some(Fill::None)));
    }
    #[test]
    fn foreground_picture_fill_preserves_use_bits_opacity_and_master_values() {
        let mut master = Paint::default();
        master.property(0x180, 3).unwrap();
        master.property(0x4186, 9).unwrap();
        master.property(0x182, 32768).unwrap();
        assert_eq!(
            Paint::default().inherit(&master).foreground_image(),
            Some((9, 32768, false))
        );

        let mut local = Paint::default();
        local.property(0x1bf, 0x00200020).unwrap();
        assert_eq!(
            local.inherit(&master).foreground_image(),
            Some((9, 32768, true))
        );
        local.property(0x1bf, 0x00100000).unwrap();
        assert!(local.inherit(&master).foreground_image().is_none());
        local.property(0x1bf, 0x00100010).unwrap();
        local.property(0x17f, 0x00010000).unwrap();
        assert!(local.inherit(&master).foreground_image().is_none());
    }

    #[test]
    fn tertiary_fill_booleans_merge_disjoint_uses_and_reject_conflicts() {
        let mut paint = Paint::default();
        paint.property(0x1bf, 0x00100010).unwrap();
        paint.tertiary_fill_boolean_property(0x00200020).unwrap();
        assert_eq!(paint.foreground_image(), None);
        paint.property(0x180, 3).unwrap();
        paint.property(0x4186, 1).unwrap();
        assert_eq!(paint.foreground_image(), Some((1, 65536, true)));

        let mut conflict = paint;
        assert!(conflict.tertiary_fill_boolean_property(0x00200000).is_err());
    }

    #[test]
    fn inactive_tertiary_bits_do_not_override_inherited_fill_rotation() {
        let mut master = Paint::default();
        master.property(0x180, 3).unwrap();
        master.property(0x4186, 2).unwrap();
        master.tertiary_fill_boolean_property(0x00200020).unwrap();
        let mut local = Paint::default();
        local.tertiary_fill_boolean_property(0x20).unwrap();
        assert_eq!(
            local.inherit(&master).foreground_image(),
            Some((2, 65536, true))
        );
    }

    #[test]
    fn foreground_picture_fill_rejects_only_unsupported_picture_placement() {
        let mut p = Paint::default();
        p.property(0x180, 3).unwrap();
        p.property(0x4186, 4).unwrap();
        assert_eq!(p.foreground_image(), Some((4, 65536, false)));

        // MS-ODRAW 2.3.7.12-13: authored fill dimensions are inert when
        // fillDztype is the default, but a nondefault sizing mode needs a
        // placement mapping beyond this plain stretch subset.
        p.property(0x189, 123).unwrap();
        p.property(0x18a, 456).unwrap();
        p.property(0x195, 0).unwrap();
        for id in 0x198..=0x19b {
            p.property(id, 0).unwrap();
        }
        assert_eq!(p.foreground_image(), Some((4, 65536, false)));
        p.property(0x19a, 1).unwrap();
        assert!(p.foreground_image().is_none());
        p.property(0x19a, 0).unwrap();
        p.property(0x195, 1).unwrap();
        assert!(p.foreground_image().is_none());
        p.property(0x195, 0).unwrap();

        // Unused Boolean values do not override their documented defaults.
        p.property(0x1bf, 0x26).unwrap();
        assert_eq!(p.foreground_image(), Some((4, 65536, false)));
        p.property(0x1bf, 0x00020002).unwrap();
        assert!(p.foreground_image().is_none());
        p.property(0x1bf, 0x00020000).unwrap();
        p.property(0x1bf, 0x00040000).unwrap();
        assert!(p.foreground_image().is_none());

        // Pattern, texture, and gradient values are distinct MSOFILLTYPEs.
        for fill_type in [1, 2, 4] {
            p.property(0x180, fill_type).unwrap();
            assert!(p.foreground_image().is_none());
        }
        p.property(0x180, 3).unwrap();
        p.property(0x4186, 0).unwrap();
        assert!(p.foreground_image().is_none());

        for id in 0x198..=0x19b {
            let mut parent = Paint::default();
            parent.property(0x180, 3).unwrap();
            parent.property(0x4186, 4).unwrap();
            parent.property(id, 1).unwrap();
            assert!(Paint::default()
                .inherit(&parent)
                .foreground_image()
                .is_none());
            let mut child = Paint::default();
            child.property(id, 0).unwrap();
            assert_eq!(
                child.inherit(&parent).foreground_image(),
                Some((4, 65536, false))
            );
        }
    }

    #[test]
    fn literal_colors_width_and_fixed_point_opacity_survive() {
        let mut p = Paint::default();
        for (id, value) in [
            (0x181, 0x00563412),
            (0x182, 32768),
            (0x1c0, 0x00efcdab),
            (0x1cb, 25400),
        ] {
            p.property(id, value).unwrap();
        }
        let (fill, stroke) = model(&p, true);
        assert_eq!(solid(&fill), Some("12345680"));
        let stroke = stroke.unwrap();
        assert_eq!((stroke.color.as_str(), stroke.width), ("ABCDEF", 25400));
        assert_eq!(stroke.line_cap.as_deref(), Some("butt"));
    }

    #[test]
    fn boolean_use_bits_control_suppression_not_the_unused_values() {
        let mut p = Paint::default();
        p.property(0x181, 0x000000ff).unwrap();
        p.property(0x1c0, 0).unwrap();
        p.property(0x1bf, 0).unwrap();
        p.property(0x1ff, 0).unwrap();
        assert_eq!(solid(&model(&p, true).0), Some("FF0000"));
        p.property(0x1bf, 0x00100000).unwrap();
        p.property(0x1ff, 0x00080000).unwrap();
        let (fill, stroke) = model(&p, true);
        assert!(matches!(fill, Some(Fill::None)));
        assert!(stroke.is_none());
    }

    #[test]
    fn does_not_invent_scheme_colors_gradient_fills_or_line_fills() {
        let mut p = Paint::default();
        p.property(0x181, 0x08000005).unwrap();
        assert_eq!(solid(&model(&p, true).0), None);
        p.property(0x181, 255).unwrap();
        p.property(0x180, 4).unwrap();
        assert_eq!(solid(&model(&p, true).0), None);
        p.property(0x180, 0).unwrap();
        assert_eq!(solid(&model(&p, true).0), Some("FF0000"));
        // A shape without a fill area (lines, connectors) gets no fill.
        assert!(matches!(model(&p, false).0, Some(Fill::None)));
    }
    #[test]
    fn validates_opacity_and_line_width_without_clamping() {
        let mut p = Paint::default();
        assert!(p.property(0x182, 65537).is_err());
        assert!(p.property(0x1c1, u32::MAX).is_err());
        assert!(p.property(0x1cb, u32::MAX).is_err());
    }

    #[test]
    fn geometry_vetoes_are_independent_and_custom_fill_rects_are_not_invented() {
        let mut p = Paint::default();
        p.property(0x181, 255).unwrap();
        p.property(0x1c0, 0).unwrap();
        p.property(0x17f, 0x00090000).unwrap();
        p.property(0x1bf, 0x00100010).unwrap();
        p.property(0x1ff, 0x00080008).unwrap();
        let (fill, stroke) = model(&p, true);
        assert!(matches!(fill, Some(Fill::None)));
        assert!(stroke.is_none());
        p.property(0x17f, 0x00090009).unwrap();
        assert_eq!(solid(&model(&p, true).0), Some("FF0000"));
        p.property(0x1bf, 0x00020002).unwrap();
        assert_eq!(solid(&model(&p, true).0), None);
    }
}
