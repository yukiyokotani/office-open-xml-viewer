//! Format-neutral OfficeArt fill and line property accumulation
//! (MS-ODRAW 2.3.6.31, 2.3.7 and 2.3.8). Hosts choose how absent properties
//! resolve (drawing defaults, masters or the normative property defaults) and
//! own every output-model adapter; this module never names an output format.
use super::unsupported;

#[derive(Clone, Copy, Default)]
pub(crate) struct Paint {
    pub(crate) details: crate::officeart::stroke::Details,
    pub(crate) custom_geometry: bool,
    /// adjustValue..adjust10Value (MS-ODRAW 2.3.6.10-19), for preset
    /// conversion by hosts that implement it (`officeart::preset`).
    pub(crate) adjust: [Option<i32>; 10],
    pub(crate) fill_ok: Option<bool>,
    pub(crate) line_ok: Option<bool>,
    pub(crate) fill_rect: Option<bool>,
    pub(crate) fill_shape: Option<bool>,
    pub(crate) rotate_fill_with_shape: Option<bool>,
    pub(crate) fill_type: Option<u32>,
    pub(crate) fill: Option<u32>,
    pub(crate) fill_blip: Option<u32>,
    pub(crate) fill_alpha: Option<u32>,
    pub(crate) fill_back: Option<u32>,
    pub(crate) fill_back_alpha: Option<u32>,
    pub(crate) fill_angle: Option<u32>,
    pub(crate) fill_focus: Option<u32>,
    pub(crate) fill_shade_type: Option<u32>,
    pub(crate) fill_dztype: Option<u32>,
    pub(crate) fill_origins: [Option<u32>; 4],
    pub(crate) filled: Option<bool>,
    pub(crate) line_type: Option<u32>,
    pub(crate) line: Option<u32>,
    pub(crate) line_alpha: Option<u32>,
    pub(crate) lined: Option<bool>,
    pub(crate) width: Option<u32>,
    pub(crate) dash: Option<u32>,
}
impl Paint {
    fn merge_boolean(
        target: &mut Option<bool>,
        value: bool,
        name: &'static str,
    ) -> Result<(), String> {
        if target.is_some_and(|current| current != value) {
            return Err(unsupported(name));
        }
        *target = Some(value);
        Ok(())
    }

    fn fill_boolean_property(&mut self, value: u32, reject_conflicts: bool) -> Result<(), String> {
        // MS-ODRAW 2.3.7.43: each low-word value is active only when its
        // corresponding high-word use bit is set. Primary and tertiary FOPT
        // tables have no documented precedence for contradictory active bits.
        if value & 0x00100000 != 0 {
            let value = value & 0x10 != 0;
            if reject_conflicts {
                Self::merge_boolean(
                    &mut self.filled,
                    value,
                    "ambiguous OfficeArt filled property",
                )?;
            } else {
                self.filled = Some(value);
            }
        }
        if value & 0x00020000 != 0 {
            let value = value & 2 != 0;
            if reject_conflicts {
                Self::merge_boolean(
                    &mut self.fill_rect,
                    value,
                    "ambiguous OfficeArt fill rectangle property",
                )?;
            } else {
                self.fill_rect = Some(value);
            }
        }
        if value & 0x00040000 != 0 {
            let value = value & 4 != 0;
            if reject_conflicts {
                Self::merge_boolean(
                    &mut self.fill_shape,
                    value,
                    "ambiguous OfficeArt fill shape property",
                )?;
            } else {
                self.fill_shape = Some(value);
            }
        }
        if value & 0x00200000 != 0 {
            let value = value & 0x20 != 0;
            if reject_conflicts {
                Self::merge_boolean(
                    &mut self.rotate_fill_with_shape,
                    value,
                    "ambiguous OfficeArt fill rotation property",
                )?;
            } else {
                self.rotate_fill_with_shape = Some(value);
            }
        }
        Ok(())
    }

    pub fn tertiary_fill_boolean_property(&mut self, value: u32) -> Result<(), String> {
        self.fill_boolean_property(value, true)
    }

    /// Explicit hspMaster supplies defaults (MS-ODRAW 1.1, 2.2.40, 2.3.2.1).
    /// Preserve absence until the full chain is resolved, especially Boolean
    /// use bits: explicit false must win over inherited true. Colors stay in
    /// their source representation until the destination slide resolves them.
    pub fn inherit(&self, parent: &Self) -> Self {
        Self {
            details: self.details.inherit(&parent.details),
            // Unsupported inherited adjustments/paths must not turn into an
            // invented unadjusted preset. Explicit paths are decoded separately.
            custom_geometry: self.custom_geometry || parent.custom_geometry,
            adjust: std::array::from_fn(|i| self.adjust[i].or(parent.adjust[i])),
            fill_ok: self.fill_ok.or(parent.fill_ok),
            line_ok: self.line_ok.or(parent.line_ok),
            fill_rect: self.fill_rect.or(parent.fill_rect),
            fill_shape: self.fill_shape.or(parent.fill_shape),
            rotate_fill_with_shape: self
                .rotate_fill_with_shape
                .or(parent.rotate_fill_with_shape),
            fill_type: self.fill_type.or(parent.fill_type),
            fill: self.fill.or(parent.fill),
            fill_blip: self.fill_blip.or(parent.fill_blip),
            fill_alpha: self.fill_alpha.or(parent.fill_alpha),
            fill_back: self.fill_back.or(parent.fill_back),
            fill_back_alpha: self.fill_back_alpha.or(parent.fill_back_alpha),
            fill_angle: self.fill_angle.or(parent.fill_angle),
            fill_focus: self.fill_focus.or(parent.fill_focus),
            fill_shade_type: self.fill_shade_type.or(parent.fill_shade_type),
            fill_dztype: self.fill_dztype.or(parent.fill_dztype),
            fill_origins: std::array::from_fn(|i| self.fill_origins[i].or(parent.fill_origins[i])),
            filled: self.filled.or(parent.filled),
            line_type: self.line_type.or(parent.line_type),
            line: self.line.or(parent.line),
            line_alpha: self.line_alpha.or(parent.line_alpha),
            lined: self.lined.or(parent.lined),
            width: self.width.or(parent.width),
            dash: self.dash.or(parent.dash),
        }
    }
    pub fn property(&mut self, id: u16, value: u32) -> Result<(), String> {
        // MS-ODRAW 2.3.6: customized vertices/segments/adjustments require their
        // own geometry conversion; never apply a preset over these overrides.
        match id {
            0x145..=0x150 => {
                if (0x147..=0x150).contains(&id) {
                    self.adjust[usize::from(id - 0x147)] = Some(value as i32);
                }
                self.custom_geometry = true;
            }
            // MS-ODRAW 2.3.6.31: geometry can veto paint independently of
            // fill/line style. Its use bits must not override those style bits.
            0x17f => {
                if value & 0x00010000 != 0 {
                    self.fill_ok = Some(value & 1 != 0);
                }
                if value & 0x00080000 != 0 {
                    self.line_ok = Some(value & 8 != 0);
                }
            }
            0x180 => self.fill_type = Some(value),
            0x181 => self.fill = Some(value),
            0x183 => self.fill_back = Some(value),
            0x184 => self.fill_back_alpha = Some(value),
            0x18b => self.fill_angle = Some(value),
            0x18c => self.fill_focus = Some(value),
            0x19c => self.fill_shade_type = Some(value),
            // Caller validates fBid on this one-based BStore reference.
            0x4186 => self.fill_blip = Some(value),
            0x182 | 0x1c1 => {
                if value > 65536 {
                    return Err(unsupported("invalid OfficeArt paint opacity"));
                }
                if id == 0x182 {
                    self.fill_alpha = Some(value);
                } else {
                    self.line_alpha = Some(value);
                }
            }
            0x195 => self.fill_dztype = Some(value),
            0x198..=0x19b => self.fill_origins[usize::from(id - 0x198)] = Some(value),
            // Boolean property's high word contains use bits, low word values.
            // MS-ODRAW 2.3.7.43 and 2.3.8.38: unused values cannot override paint.
            0x1bf => self.fill_boolean_property(value, false)?,
            0x1ff if value & 0x00080000 != 0 => self.lined = Some(value & 8 != 0),
            0x1c0 => self.line = Some(value),
            0x1c4 => self.line_type = Some(value),
            0x1cb => {
                if value > 0x132f540 {
                    return Err(unsupported("invalid OfficeArt line width"));
                }
                self.width = Some(value);
            }
            0x1ce => {
                if crate::officeart::stroke::preset_dash(value).is_none() {
                    return Err(unsupported("invalid OfficeArt line dashing"));
                }
                self.dash = Some(value);
            }
            _ => self.details.property(id, value)?,
        }
        Ok(())
    }
    pub fn geometry(&self, kind: u16) -> Option<&'static str> {
        if self.custom_geometry {
            return None;
        }
        // MS-ODRAW 2.4.24 -> ECMA-376 ST_ShapeType. Only presets whose
        // unadjusted outlines correspond directly are included here.
        match kind {
            1 | 202 => Some("rect"),
            3 => Some("ellipse"),
            4 => Some("diamond"),
            5 => Some("triangle"),
            6 => Some("rtTriangle"),
            20 => Some("line"),
            // MS-ODRAW 2.4.24: distinct from msosptLine. Preserve the
            // static preset path, not editable endpoint bindings/rerouting.
            32 => Some("straightConnector1"),
            _ => None,
        }
    }
    /// MS-ODRAW 2.3.7.1: msofillPattern (1), msofillTexture (2) and
    /// msofillPicture (3) paint with the fillBlip BLIP. Returns the active one
    /// so a caller that cannot project it can reject instead of drawing none.
    #[cfg(any(test, feature = "direct-ppt"))]
    pub fn blip_fill_type(&self) -> Option<u32> {
        let kind = self.fill_type.unwrap_or(0);
        (matches!(kind, 1..=3) && self.filled.unwrap_or(true) && self.fill_ok.unwrap_or(true))
            .then_some(kind)
    }
    pub fn background_image(&self) -> Option<(u32, u32)> {
        (self.fill_type == Some(3)
            && self.fill_blip.unwrap_or(0) != 0
            && !self.fill_rect.unwrap_or(false)
            && self.filled.unwrap_or(true)
            && self.fill_ok.unwrap_or(true))
        .then_some((
            self.fill_blip.unwrap_or(0),
            self.fill_alpha.unwrap_or(65536),
        ))
    }
    /// Plain foreground `msofillPicture` only. MS-ODRAW 2.3.7.43 defaults
    /// fillShape to 1, fillUseRect to 0, and fUseShapeAnchor to 0. A distinct
    /// fill rectangle or view-relative fill cannot be represented by this
    /// shape-local DrawingML stretch without inventing placement semantics.
    pub fn foreground_image(&self) -> Option<(u32, u32, bool)> {
        (self.fill_type == Some(3)
            && self.fill_blip.unwrap_or(0) != 0
            && !self.fill_rect.unwrap_or(false)
            && self.fill_shape.unwrap_or(true)
            && self.fill_dztype.unwrap_or(0) == 0
            && self
                .fill_origins
                .iter()
                .all(|value| value.unwrap_or(0) == 0)
            && self.filled.unwrap_or(true)
            && self.fill_ok.unwrap_or(true))
        .then_some((
            self.fill_blip.unwrap_or(0),
            self.fill_alpha.unwrap_or(65536),
            self.rotate_fill_with_shape.unwrap_or(false),
        ))
    }
    // Both output adapters consume these semantic eligibility/default rules.
    // Keep unsupported fill kinds and explicit paint vetoes out of either path.
    /// Solid fill for a host whose absent properties resolve through an
    /// unimplemented drawing-default layer: an explicit fill property is
    /// required before the MS-ODRAW defaults are used.
    pub(crate) fn solid_fill_values(&self, allow_fill: bool) -> Option<(u32, u32)> {
        let fill_set = self.fill.is_some()
            || self.filled.is_some()
            || self.fill_type.is_some()
            || self.fill_alpha.is_some();
        (allow_fill
            && fill_set
            && self.filled.unwrap_or(true)
            && self.fill_ok.unwrap_or(true)
            && !self.fill_rect.unwrap_or(false)
            && self.fill_type.unwrap_or(0) == 0)
            .then_some((
                self.fill.unwrap_or(0xffffff),
                self.fill_alpha.unwrap_or(65536),
            ))
    }

    /// Solid line counterpart of [`Self::solid_fill_values`].
    pub(crate) fn solid_line_values(&self, allow_line: bool) -> Option<(u32, u32)> {
        let line_set = self.line.is_some()
            || self.lined.is_some()
            || self.line_type.is_some()
            || self.line_alpha.is_some()
            || self.width.is_some()
            || self.dash.is_some()
            || self.details.specified();
        (allow_line
            && line_set
            && self.lined.unwrap_or(true)
            && self.line_ok.unwrap_or(true)
            && self.line_type.unwrap_or(0) == 0)
            .then_some((self.line.unwrap_or(0), self.line_alpha.unwrap_or(65536)))
    }

    /// Solid fill for a host without any drawing-default layer: every absent
    /// property takes its normative MS-ODRAW default (fFilled 1, fillColor
    /// white, fillOpacity 1.0), while explicit vetoes and non-solid fill kinds
    /// still suppress the solid projection.
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    pub(crate) fn solid_fill_or_default(&self, allow_fill: bool) -> Option<(u32, u32)> {
        (allow_fill
            && self.filled.unwrap_or(true)
            && self.fill_ok.unwrap_or(true)
            && !self.fill_rect.unwrap_or(false)
            && self.fill_type.unwrap_or(0) == 0)
            .then_some((
                self.fill.unwrap_or(0xffffff),
                self.fill_alpha.unwrap_or(65536),
            ))
    }

    /// Solid line for a host without any drawing-default layer: fLine 1,
    /// lineColor black and lineOpacity 1.0 are the MS-ODRAW defaults.
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    pub(crate) fn solid_line_or_default(&self, allow_line: bool) -> Option<(u32, u32)> {
        (allow_line
            && self.lined.unwrap_or(true)
            && self.line_ok.unwrap_or(true)
            && self.line_type.unwrap_or(0) == 0)
            .then_some((self.line.unwrap_or(0), self.line_alpha.unwrap_or(65536)))
    }
}

/// A linear OfficeArt shade projected to DrawingML stop order. Stop colours
/// stay OfficeArtCOLORREF values for the host to resolve.
pub(crate) struct LinearShade {
    pub projection: super::gradient::projection::Projection,
    /// Per-stop 16.16 opacity, parallel to `projection.stops`.
    pub alphas: Vec<u32>,
    pub scaled: bool,
    pub rotate_with_shape: bool,
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
    /// paired with its binary fill properties, and Word's own DOCX of a DOC
    /// shade agrees (msofillShade, fillAngle 0 -> `a:lin ang="5400000"
    /// scaled="0"` with the same three stops):
    /// - msofillShade (4) is `a:lin scaled="0"` and msofillShadeScale (7) is
    ///   `a:lin scaled="1"`; both use `ang` = 90 degrees - fillAngle.
    /// - fillShadeType 0 and the default 0x40000003 (gamma + sigma) produce the
    ///   same DrawingML stops; other shade types stay unsupported.
    /// - Without fillShadeColors, fillOpacity applies to every stop taken from
    ///   the fill colour and fillBackOpacity to every stop taken from the back
    ///   colour (e.g. 27525/65536 -> alpha 42000, back 0 -> alpha 0).
    ///
    /// Opacity combined with an authored shade-colour array has no evidence
    /// and is rejected, as are the path shades (5, 6) and the host-defined
    /// title shade (8).
    pub(crate) fn linear_shade(
        &self,
        source: &super::gradient::Borrowed<'_>,
        allow_fill: bool,
        work_budget: &mut usize,
        byte_budget: &mut usize,
    ) -> Result<Option<LinearShade>, String> {
        use super::gradient::projection;
        let scaled = match self.fill_type {
            Some(4) => false,
            Some(7) => true,
            Some(5 | 6 | 8) => {
                return Err(unsupported(
                    "OfficeArt path or title gradient fills are not supported yet",
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
                "OfficeArt gradient shade options are not supported yet",
            ));
        }
        let front_alpha = self.fill_alpha.unwrap_or(65_536);
        let back_alpha = self.fill_back_alpha.unwrap_or(65_536);

        // The decoded vector and projection scratch coexist. Reserve the exact
        // normalized and projected vector payloads before either allocation.
        let authored = source.decode(work_budget, byte_budget)?.unwrap_or_default();
        if !authored.is_empty() && (front_alpha != 65_536 || back_alpha != 65_536) {
            return Err(unsupported(
                "OfficeArt gradient opacity with shade colours is not supported yet",
            ));
        }
        if authored.first().is_some_and(|stop| stop.position != 0) {
            return Err(unsupported(
                "OfficeArt gradient shade colours must start at position 0",
            ));
        }
        let two_colour = authored.is_empty();
        let focus = self.fill_focus.unwrap_or(0) as i32;
        let requirements = projection::requirements(&authored, focus)?;
        let mut projection_scratch = requirements.scratch_bytes;
        let mut projection_output = requirements.output_bytes;
        let projection_bytes = projection_scratch
            .checked_add(projection_output)
            .and_then(|bytes| bytes.checked_add(requirements.output_bytes))
            .ok_or_else(|| unsupported("OfficeArt gradient byte budget overflow"))?;
        *byte_budget = byte_budget
            .checked_sub(projection_bytes)
            .ok_or_else(|| unsupported("OfficeArt gradient byte budget exceeded"))?;
        let (front, back) = if two_colour {
            (FRONT_MARKER, BACK_MARKER)
        } else {
            (
                self.fill.unwrap_or(0x00ff_ffff),
                self.fill_back.unwrap_or(0x00ff_ffff),
            )
        };
        let mut projection = projection::project(
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
            .map_err(|_| unsupported("OfficeArt gradient allocation failed"))?;
        for stop in &mut projection.stops {
            let (color, alpha) = match (two_colour, stop.color) {
                (true, FRONT_MARKER) => (self.fill.unwrap_or(0x00ff_ffff), front_alpha),
                (true, _) => (self.fill_back.unwrap_or(0x00ff_ffff), back_alpha),
                (false, color) => (color, 65_536),
            };
            stop.color = color;
            alphas.push(alpha);
        }
        Ok(Some(LinearShade {
            projection,
            alphas,
            scaled,
            rotate_with_shape: self.rotate_fill_with_shape.unwrap_or(false),
        }))
    }
}
