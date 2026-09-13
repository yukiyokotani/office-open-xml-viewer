//! Bounded DOC table-style character and paragraph formatting. [MS-DOC]
//! 2.4.6.6 and 2.9.41 define conditional order and CNFOperand framing;
//! 2.9.340 requires sprmTIstd inside UpxTapx to be ignored.

use super::cnf;
use super::{Formatting, Properties, Sprms, MAX_TABLE_AWARE_CACHE_ENTRIES};
use crate::doc::{paragraph, sprm, u16_at, unsupported};
use std::collections::BTreeMap;
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

#[derive(Clone)]
pub(super) struct Profile {
    unconditional: Properties,
    conditional: BTreeMap<u16, Properties>,
    pub(super) conditional_paragraph_alignment: BTreeMap<u16, paragraph::AlignmentPatch>,
    condition_presence: u16,
    bands: Bands,
    pub(super) paragraph_alignment: Option<paragraph::AlignmentPatch>,
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
        let mut has_conditional_character = false;
        let mut has_conditional_paragraph = false;
        for style_id in chain {
            let sets = *self.styles[style_id]
                .as_ref()
                .expect("validated style")
                .table
                .as_ref()
                .expect("validated table style");

            let mut tapx = Sprms::new(sets.tapx);
            while let Some((code, operand)) = tapx.next(&mut self.budget)? {
                match code {
                    0x3488 | 0x3489 => {
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
                    }
                    0x563a => {
                        // Required ignored value inside UpxTapx, MS-DOC 2.9.340.
                    }
                    _ => profile.unsupported_table = true,
                }
            }

            let mut chpx = Sprms::new(sets.chpx);
            while let Some((code, operand)) = chpx.next(&mut self.budget)? {
                match code {
                    0x2a42 | 0x4a43 | 0x4a4f | 0x4a51 | 0x6870 => {
                        let baseline = profile.unconditional.clone();
                        profile.unconditional.apply(code, operand, &baseline)?;
                    }
                    0xca85 => {
                        has_conditional_character = true;
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
                            has_conditional_paragraph = true;
                            parse_conditional_paragraph(&mut profile, operand, budget)?;
                        } else {
                            profile.unsupported_paragraph = true;
                        }
                        Ok(())
                    },
                )?;
            }
        }
        if inherited && has_conditional_character {
            // Conditional inheritance priority is not established by the
            // unconditional table-color controls.
            profile.unsupported_character = true;
        }
        if inherited && has_conditional_paragraph {
            // Conditional PAPX inheritance priority has not been established
            // by the single-style Office controls.
            profile.unsupported_paragraph = true;
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
