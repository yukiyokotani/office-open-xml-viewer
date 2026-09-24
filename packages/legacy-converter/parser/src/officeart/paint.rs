//! Format-neutral OfficeArt fill and line property accumulation
//! (MS-ODRAW 2.3.6.31, 2.3.7 and 2.3.8). Hosts choose how absent properties
//! resolve (drawing defaults, masters or the normative property defaults) and
//! own every output-model adapter; this module never names an output format.
use super::unsupported;

#[derive(Clone, Copy, Default)]
pub(crate) struct Paint {
    pub(crate) details: crate::officeart::stroke::Details,
    pub(crate) custom_geometry: bool,
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
            0x145..=0x150 => self.custom_geometry = true,
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
