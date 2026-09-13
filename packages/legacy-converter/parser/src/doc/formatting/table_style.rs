//! Bounded DOC table-style character and paragraph formatting. [MS-DOC]
//! 2.4.6.6 and 2.9.41 define conditional order and CNFOperand framing;
//! 2.9.340 requires sprmTIstd inside UpxTapx to be ignored.

use super::{cnf, tapx};
use super::{Formatting, Properties, Sprms, MAX_TABLE_AWARE_CACHE_ENTRIES};
use crate::doc::{paragraph, sprm, table, u16_at, unsupported};
use std::collections::{BTreeMap, BTreeSet};
use std::rc::Rc;

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(in crate::doc) struct TableFormattingKey {
    pub(super) selected_style: usize,
    pub(super) matches: [Option<u16>; 5],
}

impl TableFormattingKey {
    #[cfg(test)]
    pub(super) fn unconditional(selected_style: usize) -> Self {
        Self {
            selected_style,
            matches: [None; 5],
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct Bands {
    pub(super) horizontal: Option<u8>,
    pub(super) vertical: Option<u8>,
}

#[cfg(feature = "direct-doc")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) struct ConditionalTableBorders {
    pub(in crate::doc) condition: u16,
    pub(in crate::doc) sides: [Option<table::PreparedBorder>; 6],
}

#[derive(Clone)]
enum TableStyleShading {
    /// Authored ShdNil is retained as property presence even though its value
    /// is ignored when the selected style is finally applied.
    Nil,
    Value(table::Shading),
}

#[derive(Clone)]
pub(super) struct Profile {
    unconditional: Properties,
    conditional: BTreeMap<u16, Properties>,
    pub(super) conditional_paragraph_alignment: BTreeMap<u16, paragraph::AlignmentPatch>,
    condition_presence: u16,
    bands: Bands,
    pub(super) paragraph_alignment: Option<paragraph::AlignmentPatch>,
    table_shading: Option<TableStyleShading>,
    table_default_margins: table::MarginPatch,
    table_style_margins: table::MarginPatch,
    #[cfg(feature = "direct-doc")]
    table_borders: [Option<table::PreparedBorder>; 6],
    #[cfg(feature = "direct-doc")]
    conditional_table_borders: Option<ConditionalTableBorders>,
    conditional_table_shading: BTreeMap<u16, table::Shading>,
    conditional_table_shading_nil: BTreeSet<u16>,
    unsupported_character: bool,
    unsupported_paragraph: bool,
    unsupported_table: bool,
}

impl Default for Profile {
    fn default() -> Self {
        Self {
            unconditional: Properties::sparse(),
            conditional: BTreeMap::new(),
            conditional_paragraph_alignment: BTreeMap::new(),
            condition_presence: 0,
            bands: Bands::default(),
            paragraph_alignment: None,
            table_shading: None,
            table_default_margins: table::MarginPatch::default(),
            table_style_margins: table::MarginPatch::default(),
            #[cfg(feature = "direct-doc")]
            table_borders: [None; 6],
            #[cfg(feature = "direct-doc")]
            conditional_table_borders: None,
            conditional_table_shading: BTreeMap::new(),
            conditional_table_shading_nil: BTreeSet::new(),
            unsupported_character: false,
            unsupported_paragraph: false,
            unsupported_table: false,
        }
    }
}

impl Formatting<'_> {
    pub(in crate::doc) fn table_style_selector_profile(
        &mut self,
        selected_style: Option<usize>,
    ) -> Result<(Option<u8>, Option<u8>, u16), String> {
        let Some(selected_style) = selected_style else {
            return Ok((None, None, 0));
        };
        let profile = self.table_style_profile(selected_style)?;
        Ok((
            profile.bands.horizontal,
            profile.bands.vertical,
            profile.condition_presence,
        ))
    }

    pub(in crate::doc) fn table_formatting_key(
        &mut self,
        selected_style: Option<usize>,
        table_style_options: Option<u16>,
        matches: [Option<u16>; 5],
    ) -> Result<Option<TableFormattingKey>, String> {
        let Some(selected_style) = selected_style else {
            return Ok(None);
        };
        let _ = self.table_style_profile(selected_style)?;
        if table_style_options.is_none() && matches.iter().any(Option::is_some) {
            return Err(unsupported(
                "Word table style conditions selected without sprmTTlp",
            ));
        }
        for condition in matches.into_iter().flatten() {
            if !cnf::CONDITIONS.contains(&condition) {
                return Err(unsupported("invalid Word table style condition key"));
            }
        }
        Ok(Some(TableFormattingKey {
            selected_style,
            matches,
        }))
    }

    pub(super) fn apply_table_character_style(
        &mut self,
        props: &mut Properties,
        key: Option<TableFormattingKey>,
    ) -> Result<(), String> {
        let Some(key) = key else {
            return Ok(());
        };
        let profile = self.table_style_profile(key.selected_style)?;
        props.overlay_visible(&profile.unconditional);
        for condition in key.matches.into_iter().flatten() {
            if let Some(patch) = profile.conditional.get(&condition) {
                props.overlay_visible(patch);
            }
        }
        Ok(())
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn table_cell_shading(
        &mut self,
        key: Option<TableFormattingKey>,
    ) -> Result<Option<table::Shading>, String> {
        let Some(key) = key else {
            return Ok(None);
        };
        let profile = self.table_style_profile(key.selected_style)?;
        let mut shading = match &profile.table_shading {
            Some(TableStyleShading::Value(value)) => Some(value.clone()),
            Some(TableStyleShading::Nil) | None => None,
        };
        for condition in key.matches.into_iter().flatten() {
            if profile.conditional_table_shading_nil.contains(&condition) {
                // [MS-DOC] 2.9.247: a selected conditional ShdNil is
                // present, but does not affect the current shading value.
                continue;
            }
            if let Some(patch) = profile.conditional_table_shading.get(&condition) {
                shading = Some(patch.clone());
            }
        }
        Ok(shading)
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn table_cell_margins(
        &mut self,
        selected_style: Option<usize>,
    ) -> Result<(table::MarginPatch, table::MarginPatch), String> {
        let Some(selected_style) = selected_style else {
            return Ok(Default::default());
        };
        let profile = self.table_style_profile(selected_style)?;
        Ok((profile.table_default_margins, profile.table_style_margins))
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn table_borders(
        &mut self,
        selected_style: Option<usize>,
    ) -> Result<[Option<table::PreparedBorder>; 6], String> {
        let Some(selected_style) = selected_style else {
            return Ok([None; 6]);
        };
        Ok(self.table_style_profile(selected_style)?.table_borders)
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn conditional_table_borders(
        &mut self,
        selected_style: Option<usize>,
    ) -> Result<Option<(u16, [Option<table::PreparedBorder>; 6], u16)>, String> {
        let Some(selected_style) = selected_style else {
            return Ok(None);
        };
        let profile = self.table_style_profile(selected_style)?;
        Ok(profile
            .conditional_table_borders
            .map(|patch| (patch.condition, patch.sides, profile.condition_presence)))
    }

    pub(super) fn table_style_profile(&mut self, id: usize) -> Result<Rc<Profile>, String> {
        let profile = if let Some(profile) = self.table_style_cache.get(&id) {
            Rc::clone(profile)
        } else {
            let profile = Rc::new(self.build_table_style_profile(id)?);
            if self.table_style_cache.len() >= MAX_TABLE_AWARE_CACHE_ENTRIES {
                self.table_style_cache.clear();
            }
            self.table_style_cache.insert(id, Rc::clone(&profile));
            profile
        };
        self.unsupported_character_properties |= profile.unsupported_character;
        self.unsupported_paragraph_properties |= profile.unsupported_paragraph;
        self.unsupported_table_properties |= profile.unsupported_table;
        Ok(profile)
    }

    fn build_table_style_profile(&mut self, id: usize) -> Result<Profile, String> {
        if self
            .styles
            .get(id)
            .and_then(Option::as_ref)
            .filter(|style| style.kind == 3)
            .is_none()
        {
            return Ok(Profile {
                unsupported_table: true,
                ..Profile::default()
            });
        }

        let chain = self.chain(id, 3)?;
        let inherited = chain.len() > 1;
        let mut profile = Profile::default();
        let mut horizontal_source = None;
        let mut vertical_source = None;
        let mut has_conditional_table = false;
        #[cfg(feature = "direct-doc")]
        let mut conditional_border_rejected = false;
        let mut default_margin_sides = 0u8;
        let mut style_margin_sides = 0u8;
        let interpret_table_styles = self.interpret_table_styles;
        for style_id in chain {
            let sets = *self.styles[style_id]
                .as_ref()
                .expect("validated style")
                .table
                .as_ref()
                .expect("validated table style");

            let report = tapx::validate(
                style_id,
                sets.tapx,
                &mut self.budget,
                |scope, code, operand, _| {
                    match code {
                        #[cfg(feature = "direct-doc")]
                        0xd47f | 0xd680 | 0xd681 | 0xd682 | 0xd683 | 0xd684 => {
                            let tapx::Scope::Conditional(condition) = scope else {
                                return Ok(false);
                            };
                            if operand.len() != 9 || operand[0] != 8 {
                                return Err(unsupported(
                                    "invalid Word conditional table-style border",
                                ));
                            }
                            let value = table::PreparedBorder::read(&operand[1..], false)?;
                            if value.is_nil()
                                || !matches!(
                                    condition,
                                    crate::doc::table_style_condition::FIRST_ROW
                                        | crate::doc::table_style_condition::LAST_ROW
                                        | crate::doc::table_style_condition::FIRST_COLUMN
                                        | crate::doc::table_style_condition::LAST_COLUMN
                                )
                            {
                                conditional_border_rejected = true;
                                profile.conditional_table_borders = None;
                                return Ok(false);
                            }
                            let side = match code {
                                0xd47f => 0,
                                0xd680 => 2,
                                0xd681 => 1,
                                0xd682 => 3,
                                0xd683 => 4,
                                0xd684 => 5,
                                _ => unreachable!(),
                            };
                            if conditional_border_rejected {
                                return Ok(false);
                            }
                            let mut patch = match profile.conditional_table_borders {
                                Some(patch) if patch.condition != condition => {
                                    conditional_border_rejected = true;
                                    profile.conditional_table_borders = None;
                                    return Ok(false);
                                }
                                Some(patch) => patch,
                                None => ConditionalTableBorders {
                                    condition,
                                    sides: [None; 6],
                                },
                            };
                            patch.sides[side] = Some(value);
                            profile.conditional_table_borders = Some(patch);
                            profile.condition_presence |= condition;
                            has_conditional_table = true;
                            Ok(interpret_table_styles)
                        }
                        #[cfg(feature = "direct-doc")]
                        0xd613 => {
                            if scope != tapx::Scope::Unconditional {
                                return Ok(false);
                            }
                            if operand.len() != 49 || operand[0] != 48 {
                                return Err(unsupported("invalid Word table-style border array"));
                            }
                            let mut borders = [None; 6];
                            for (side, slot) in borders.iter_mut().enumerate() {
                                *slot = Some(table::PreparedBorder::read(
                                    &operand[1 + side * 8..],
                                    false,
                                )?);
                            }
                            if borders
                                .iter()
                                .flatten()
                                .copied()
                                .any(table::PreparedBorder::is_nil)
                            {
                                // NilBrc has not been established for the
                                // inherited native table-border cascade.
                                return Ok(false);
                            }
                            profile.table_borders = borders;
                            Ok(interpret_table_styles)
                        }
                        0xd634 | 0xd63e => {
                            if scope != tapx::Scope::Unconditional {
                                return Ok(false);
                            }
                            let (patch, own_sides, other_sides) = if code == 0xd634 {
                                (
                                    &mut profile.table_default_margins,
                                    &mut default_margin_sides,
                                    style_margin_sides,
                                )
                            } else {
                                (
                                    &mut profile.table_style_margins,
                                    &mut style_margin_sides,
                                    default_margin_sides,
                                )
                            };
                            let sides = patch.apply_style(code, operand)?;
                            // The controls establish D63E above direct D634,
                            // and each property independently through basedOn.
                            // They do not establish cross-property composition
                            // within one style chain on the same side.
                            if sides & other_sides != 0 {
                                profile.unsupported_table = true;
                            }
                            *own_sides |= sides;
                            Ok(interpret_table_styles)
                        }
                        0xd687 => {
                            if operand.len() != 11 || operand[0] != 10 {
                                return Err(unsupported(
                                    "invalid Word table-style shading operand",
                                ));
                            }
                            let bytes = &operand[1..];
                            if table::Shading::is_shd_nil(bytes) {
                                match scope {
                                    tapx::Scope::Unconditional => {
                                        // [MS-DOC] 2.9.247 ignores ShdNil when
                                        // the selected style is applied. Keep
                                        // its authored presence distinct from
                                        // an omitted property: Word controls
                                        // show that an inherited child Nil does
                                        // not behave like an empty child.
                                        profile.table_shading = Some(TableStyleShading::Nil);
                                        return Ok(interpret_table_styles);
                                    }
                                    tapx::Scope::Conditional(condition) => {
                                        // A conditional ShdNil is an authored
                                        // property which applies as a no-op. It
                                        // replaces an inherited patch for the
                                        // same condition without clearing the
                                        // unconditional or earlier condition.
                                        profile.condition_presence |= condition;
                                        profile.conditional_table_shading.remove(&condition);
                                        profile.conditional_table_shading_nil.insert(condition);
                                        has_conditional_table = true;
                                        return Ok(interpret_table_styles);
                                    }
                                }
                            }
                            let Some(shading) = table::Shading::read(bytes, false)? else {
                                return Ok(false);
                            };
                            match scope {
                                tapx::Scope::Unconditional => {
                                    profile.table_shading = Some(TableStyleShading::Value(shading));
                                }
                                tapx::Scope::Conditional(condition) => {
                                    has_conditional_table = true;
                                    profile.condition_presence |= condition;
                                    profile.conditional_table_shading_nil.remove(&condition);
                                    profile.conditional_table_shading.insert(condition, shading);
                                }
                            }
                            // Retain exact facts for both paths, but only the
                            // direct model currently projects table-style cell
                            // shading. The XML conversion keeps its admission
                            // gate rather than silently omitting this property.
                            Ok(interpret_table_styles)
                        }
                        0x3488 | 0x3489 => {
                            if scope != tapx::Scope::Unconditional {
                                return Ok(false);
                            }
                            let value = *operand
                                .first()
                                .ok_or_else(|| unsupported("short Word table style band size"))?;
                            if !(1..=3).contains(&value) {
                                return Err(unsupported("invalid Word table style band size"));
                            }
                            let (slot, source) = if code == 0x3488 {
                                (&mut profile.bands.horizontal, &mut horizontal_source)
                            } else {
                                (&mut profile.bands.vertical, &mut vertical_source)
                            };
                            if slot.is_some_and(|prior| prior != value)
                                && source.is_some_and(|prior| prior != style_id)
                            {
                                // UpxTapx inheritance order is not inferred from
                                // unconditional CHPX evidence.
                                profile.unsupported_table = true;
                            }
                            *slot = Some(value);
                            *source = Some(style_id);
                            Ok(true)
                        }
                        // The validator permits only the required zero dxa value
                        // in default style 0x000B. It introduces no leading indent.
                        0xf617 if scope == tapx::Scope::Unconditional => Ok(true),
                        _ => Ok(false),
                    }
                },
            )?;
            profile.unsupported_table |= report.unsupported;

            let mut chpx = Sprms::new(sets.chpx);
            while let Some((code, operand)) = chpx.next(&mut self.budget)? {
                match code {
                    0x2a42 | 0x4a43 | 0x4a4f | 0x4a51 | 0x6870 => {
                        let baseline = profile.unconditional.clone();
                        profile.unconditional.apply(code, operand, &baseline)?;
                    }
                    0xca85 => {
                        parse_conditional(&mut profile, operand, &mut self.budget)?;
                    }
                    _ => profile.unsupported_character = true,
                }
            }

            if !sets.papx.is_empty() {
                let embedded_style = usize::from(u16_at(sets.papx, 0)?);
                if embedded_style != style_id {
                    // MS-DOC 2.9.338 UpxPapx requires this optional istd,
                    // when present, to equal the current style.
                    return Err(unsupported(
                        "Word table style PAPX has mismatched style index",
                    ));
                }
                sprm::paragraph_properties(
                    &sets.papx[2..],
                    self.data,
                    &mut self.budget,
                    |code, operand, budget| {
                        if code == 0x2461 {
                            let alignment = paragraph::AlignmentPatch::from_sprm(code, operand)?
                                .expect("logical alignment code");
                            profile.paragraph_alignment = Some(alignment);
                        } else if code == 0x2403 {
                            let _ = paragraph::AlignmentPatch::from_sprm(code, operand)?;
                            // Office 16.112.4 table-style controls ignore
                            // physical PJc80. Keep direct paragraph PJc80
                            // behavior independent and retain admission gating.
                            profile.unsupported_paragraph = true;
                        } else if code == 0xc666 {
                            parse_conditional_paragraph(&mut profile, operand, budget)?;
                        } else {
                            profile.unsupported_paragraph = true;
                        }
                        Ok(())
                    },
                )?;
            }
        }
        // Native Word controls establish field-wise base-to-child composition
        // for supported conditional color, absolute CHps and logical PJc,
        // including empty and nonmatching children and direct overrides. The
        // selected conditions still apply after unconditional properties. This
        // evidence does not cover conditional fonts, PJc80 or other properties.
        if inherited && has_conditional_table {
            // Native controls have not yet established inherited TCnf
            // composition or priority. Keep it visible to the admission gate.
            profile.unsupported_table = true;
            #[cfg(feature = "direct-doc")]
            {
                profile.conditional_table_borders = None;
            }
        }
        Ok(profile)
    }
}

fn parse_conditional_paragraph(
    profile: &mut Profile,
    operand: &[u8],
    budget: &mut super::Budget,
) -> Result<(), String> {
    let operand = cnf::parse(operand)?;
    let condition = operand.condition;
    let mut patch = profile
        .conditional_paragraph_alignment
        .get(&condition)
        .copied();
    let mut has_supported_alignment = false;
    let mut nested = Sprms::new(operand.grpprl);
    while let Some((code, value)) = nested.next(budget)? {
        if code == 0x2461 {
            patch = paragraph::AlignmentPatch::from_sprm(code, value)?;
            has_supported_alignment = true;
        } else {
            // This includes PJc80 and nested CNF records. The bounded Office
            // evidence establishes conditional PJc only; unsupported records
            // do not contribute condition presence.
            if code == 0x2403 {
                let _ = paragraph::AlignmentPatch::from_sprm(code, value)?;
            }
            profile.unsupported_paragraph = true;
        }
    }
    if has_supported_alignment {
        profile.condition_presence |= condition;
        profile
            .conditional_paragraph_alignment
            .insert(condition, patch.expect("supported alignment"));
    }
    Ok(())
}

fn parse_conditional(
    profile: &mut Profile,
    operand: &[u8],
    budget: &mut super::Budget,
) -> Result<(), String> {
    let operand = cnf::parse(operand)?;
    let condition = operand.condition;
    let mut patch = profile
        .conditional
        .get(&condition)
        .cloned()
        .unwrap_or_else(Properties::sparse);
    let mut has_supported_character = false;
    let mut nested = Sprms::new(operand.grpprl);
    while let Some((code, value)) = nested.next(budget)? {
        if matches!(code, 0x2a42 | 0x4a43 | 0x6870) {
            let baseline = patch.clone();
            patch.apply(code, value, &baseline)?;
            has_supported_character = true;
        } else if matches!(code, 0x4a4f | 0x4a50 | 0x4a51 | 0x4a5e) {
            // Word 16.112.4 controls with seven reordered FFN records leave
            // conditional font markers fixed while unconditional CRgFtc values
            // follow the reordered records. That rules out treating these as a
            // simple local font-table index, but does not establish a remap.
            profile.unsupported_character = true;
        } else {
            // This includes nested CNF records. They are parsed only as one
            // bounded operand and are never recursively expanded.
            profile.unsupported_character = true;
        }
    }
    if has_supported_character {
        // Office 16.112.4 controls show that an empty CCnf has no conditional
        // presence, while a supported color or absolute-size property does.
        // Unsupported property families remain gated and do not broaden it.
        profile.condition_presence |= condition;
        profile.conditional.insert(condition, patch);
    }
    Ok(())
}
