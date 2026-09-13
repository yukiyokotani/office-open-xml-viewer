//! Bounded DOC table-style character formatting. [MS-DOC] 2.4.6.6 and
//! 2.9.41 define conditional order and CNFOperand framing; 2.9.340 requires
//! sprmTIstd inside UpxTapx to be ignored.

use super::{Formatting, Properties, Sprms, MAX_TABLE_AWARE_CACHE_ENTRIES};
use crate::doc::{table_style_condition, u16_at, unsupported};
use std::collections::BTreeMap;
use std::rc::Rc;

const CONDITIONS: [u16; 12] = [
    0x0001, 0x0002, 0x0004, 0x0008, 0x0010, 0x0020, 0x0040, 0x0080, 0x0100, 0x0200, 0x0400, 0x0800,
];

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
    condition_presence: u16,
    bands: Bands,
    unsupported_character: bool,
    unsupported_table: bool,
}

impl Default for Profile {
    fn default() -> Self {
        Self {
            unconditional: Properties::sparse(),
            conditional: BTreeMap::new(),
            condition_presence: 0,
            bands: Bands::default(),
            unsupported_character: false,
            unsupported_table: false,
        }
    }
}

impl Formatting<'_> {
    pub(in crate::doc) fn table_style_bands(
        &mut self,
        selected_style: Option<usize>,
    ) -> Result<(Option<u8>, Option<u8>), String> {
        let Some(selected_style) = selected_style else {
            return Ok((None, None));
        };
        let bands = self.table_style_profile(selected_style)?.bands;
        Ok((bands.horizontal, bands.vertical))
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
        let profile = self.table_style_profile(selected_style)?;
        if let Some(options) = table_style_options {
            // Controlled Word output shows that an empty first-column CCnf
            // does not shift band numbering, while the tested nonempty CCnf
            // bold and PCnf alignment records do. Supported color CCnf is
            // resolved here; other edge/band interactions remain gated.
            let horizontal_enabled = options & (1 << 9) == 0 && profile.bands.horizontal.is_some();
            let vertical_enabled = options & (1 << 10) == 0 && profile.bands.vertical.is_some();
            let unresolved_horizontal = horizontal_enabled
                && ((options & (1 << 5) != 0
                    && profile.condition_presence & table_style_condition::FIRST_ROW == 0)
                    || (options & (1 << 6) != 0
                        && profile.condition_presence & table_style_condition::LAST_ROW == 0));
            let unresolved_vertical = vertical_enabled
                && ((options & (1 << 7) != 0
                    && profile.condition_presence & table_style_condition::FIRST_COLUMN == 0)
                    || (options & (1 << 8) != 0
                        && profile.condition_presence & table_style_condition::LAST_COLUMN == 0));
            if unresolved_horizontal || unresolved_vertical {
                self.unsupported_character_properties = true;
            }
        } else if matches.iter().any(Option::is_some) {
            return Err(unsupported(
                "Word table style conditions selected without sprmTTlp",
            ));
        }
        for condition in matches.into_iter().flatten() {
            if !CONDITIONS.contains(&condition) {
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
        let mut has_conditional = false;
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
                    0x2a42 | 0x6870 => {
                        let baseline = profile.unconditional.clone();
                        profile.unconditional.apply(code, operand, &baseline)?;
                    }
                    0xca85 => {
                        has_conditional = true;
                        parse_conditional(&mut profile, operand, &mut self.budget)?;
                    }
                    _ => profile.unsupported_character = true,
                }
            }
        }
        if inherited && has_conditional {
            // Conditional inheritance priority is not established by the
            // unconditional table-color controls.
            profile.unsupported_character = true;
        }
        Ok(profile)
    }
}

fn parse_conditional(
    profile: &mut Profile,
    operand: &[u8],
    budget: &mut super::Budget,
) -> Result<(), String> {
    if operand.len() < 3 || usize::from(operand[0]) + 1 != operand.len() {
        return Err(unsupported("invalid Word conditional formatting operand"));
    }
    let condition = u16_at(operand, 1)?;
    if !CONDITIONS.contains(&condition) {
        return Err(unsupported("invalid Word table style condition"));
    }
    let mut patch = profile
        .conditional
        .get(&condition)
        .cloned()
        .unwrap_or_else(Properties::sparse);
    let mut has_supported_color = false;
    let mut nested = Sprms::new(&operand[3..]);
    while let Some((code, value)) = nested.next(budget)? {
        if matches!(code, 0x2a42 | 0x6870) {
            let baseline = patch.clone();
            patch.apply(code, value, &baseline)?;
            has_supported_color = true;
        } else {
            // This includes nested CNF records. They are parsed only as one
            // bounded operand and are never recursively expanded.
            profile.unsupported_character = true;
        }
    }
    if has_supported_color {
        // Office 16.112.4 controls show that an empty first-column CCnf does
        // not shift band numbering, while first-column PCnf alignment and
        // CCnf bold true/false do. Those properties remain unsupported here;
        // track only a color contribution this slice can apply and verify.
        profile.condition_presence |= condition;
        profile.conditional.insert(condition, patch);
    }
    Ok(())
}
