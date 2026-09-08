//! Typed list-marker projection for the direct DOC model path.
//!
//! [MS-DOC] 2.4.6.3–2.4.6.4: counters are shared by LSID while start
//! overrides apply once to an LFO/level instance. LVLF.rgbxchNums is the sole
//! authority for placeholders; literal `%N` text is never reinterpreted.

use super::{Level, Reference, Tables};
use crate::doc::{character, number_format, unsupported};
use docx_model::{DocParagraph, NumberingInfo};
use std::collections::{HashMap, HashSet};

#[derive(Clone)]
struct ListState {
    // Triangular (level, restart-threshold) summaries: 1+2+...+9 = 45.
    virtual_values: [VirtualCounter; 45],
}

#[derive(Clone, Copy)]
struct AncestorValue {
    value: Option<u32>,
    format: Option<&'static str>,
}

#[derive(Clone, Copy, Default)]
enum VirtualCounter {
    #[default]
    Empty,
    /// Number of real paragraphs at this level since the applicable restart.
    /// [MS-DOC] 2.4.6.4 initializes each query from the current LVL's
    /// `iStartAt`, so this must not retain an earlier LVL's start value.
    Relative(u32),
    /// Last absolute value established by a once-only LFO start override.
    /// Subsequent occurrences advance that value independently of LVL starts.
    Absolute(u32),
}

impl Default for ListState {
    fn default() -> Self {
        Self {
            virtual_values: [VirtualCounter::Empty; 45],
        }
    }
}

fn summary_index(level: u8, threshold: u8) -> usize {
    usize::from(level) * (usize::from(level) + 1) / 2 + usize::from(threshold)
}

#[derive(Default)]
pub struct Store {
    scope: usize,
    lists: HashMap<i32, ListState>,
    started: HashSet<(usize, u8)>,
    // Ancestor placeholders belong to a DOC "list" (same iLfo), unlike the
    // numeric sequences above, which MS-DOC 2.4.6.4 shares by LSID.
    ancestors: HashMap<(usize, u8), AncestorValue>,
}

impl Store {
    pub fn begin_story(&mut self) -> Result<(), String> {
        self.scope = self
            .scope
            .checked_add(1)
            .ok_or_else(|| unsupported("Word numbering story counter overflow"))?;
        // Stories never resume after projection. Drop their LSID/LFO state so
        // headers, notes, and the body cannot accumulate counter maps for the
        // lifetime of the source archive.
        self.lists.clear();
        self.started.clear();
        self.ancestors.clear();
        Ok(())
    }

    pub fn activate(
        &mut self,
        tables: &Tables<'_>,
        reference: Reference,
        marker: &character::Properties,
        paragraph: &DocParagraph,
        fonts: &[String],
    ) -> Result<NumberingInfo, String> {
        let selection = tables.resolve(reference)?;
        if selection.instance.auto_number_field.is_some() {
            return Err(unsupported(
                "direct DOC model does not execute automatic-number fields",
            ));
        }
        let effective = |index: u8| -> Option<&Level<'_>> {
            selection
                .instance
                .levels
                .iter()
                .find(|value| value.index == index)
                .and_then(|value| value.formatting.as_ref())
                .or_else(|| selection.list.levels.get(usize::from(index)))
        };
        let current =
            effective(reference.level).ok_or_else(|| unsupported("missing Word list level"))?;
        // LVLF.fLegal is defined as overriding inherited level-number formats,
        // while 2.4.6.3 Part 2 step 4 can be read as reformatting every
        // placeholder. Without a Word control for the current non-Arabic
        // level, reject only that ambiguous combination.
        if current.legal && !matches!(current.format, 0 | 0x16) {
            return Err(unsupported(
                "legal Word list level with a non-Arabic current format",
            ));
        }
        let template = current.numbering_template(reference.level, |index| {
            effective(index).map(|level| effective_format(current, level, index < reference.level))
        })?;
        let mut formats = ["decimal"; 9];
        for index in 0..=reference.level {
            if let Some(level) = effective(index) {
                formats[usize::from(index)] =
                    number_format::name(effective_format(current, level, index < reference.level))?;
            }
        }
        // Validate marker font references before mutating counter state. Later
        // counter/output failures are fatal to the atomic direct constructor.
        let facts = marker.direct_font_facts(fonts)?;
        let mut state = self
            .lists
            .get(&selection.list.id)
            .cloned()
            .unwrap_or_default();
        let numberless = matches!(current.format, 0x17 | 0xff) || current.start.is_none();
        let start = current.start.map(u32::from);
        let first_instance = !self.started.contains(&(reference.index, reference.level));
        let absolute = first_instance
            .then_some(selection.start_override)
            .flatten()
            .map(u32::from);
        // MS-DOC 2.4.6.4 normatively counts prior numberless paragraphs, but
        // does not settle the interaction between step 2's early return and
        // the first-use timing adaptation established for numeric LFO start
        // overrides. Reject this precise unmeasured combination.
        if numberless && absolute.is_some() {
            return Err(unsupported(
                "numberless Word list level with a start override",
            ));
        }
        let mut selected = None;
        // A bullet/none paragraph has no number for its own query (step 2),
        // but it remains a real same-LSID/level paragraph when a later query
        // scans prior paragraphs in steps 6-11. Advance every policy summary
        // and record its first LFO encounter even when this marker is numberless.
        for threshold in 0..=reference.level {
            let index = summary_index(reference.level, threshold);
            let (value, next) = match (absolute, state.virtual_values[index]) {
                (Some(value), _) => (Some(value), VirtualCounter::Absolute(value)),
                (None, VirtualCounter::Empty) => (start, VirtualCounter::Relative(1)),
                (None, VirtualCounter::Relative(prior)) => {
                    let value = start
                        .map(|start| {
                            start.checked_add(prior).ok_or_else(|| {
                                counter_error(ooxml_common::numbering::CounterError::Overflow)
                            })
                        })
                        .transpose()?;
                    let count = prior.checked_add(1).ok_or_else(|| {
                        counter_error(ooxml_common::numbering::CounterError::Overflow)
                    })?;
                    (value, VirtualCounter::Relative(count))
                }
                (None, VirtualCounter::Absolute(prior)) => {
                    let value = prior.checked_add(1).ok_or_else(|| {
                        counter_error(ooxml_common::numbering::CounterError::Overflow)
                    })?;
                    (Some(value), VirtualCounter::Absolute(value))
                }
            };
            state.virtual_values[index] = next;
            if current.restart == Some(threshold) {
                selected = value;
            }
        }
        let counter = if numberless {
            0
        } else {
            let value = selected.ok_or_else(|| unsupported("invalid Word list restart limit"))?;
            value
        };
        // This real paragraph restarts every deeper virtual sequence whose
        // candidate threshold is greater than the encountered level.
        for deeper in reference.level + 1..=8 {
            for threshold in reference.level + 1..=deeper {
                state.virtual_values[summary_index(deeper, threshold)] = VirtualCounter::Empty;
            }
        }
        let mut ancestor_values = [None; 9];
        for target in current
            .placeholders
            .iter()
            .flatten()
            .map(|(_, target)| *target)
        {
            if target >= reference.level {
                continue;
            }
            let ancestor = self
                .ancestors
                .get(&(reference.index, target))
                .ok_or_else(|| unsupported("missing prior Word list ancestor level"))?;
            if matches!(
                effective(target).map(|level| level.format),
                Some(0x17 | 0xff)
            ) {
                // NumberingTemplate omits this placeholder; the explicit
                // ancestor record distinguishes an encountered empty level
                // number from a missing prior paragraph.
                continue;
            }
            ancestor_values[usize::from(target)] = Some(
                ancestor
                    .value
                    .ok_or_else(|| unsupported("incompatible Word list ancestor level"))?,
            );
            if !current.legal {
                formats[usize::from(target)] = ancestor
                    .format
                    .ok_or_else(|| unsupported("incompatible Word list ancestor format"))?;
            }
        }
        let text = template
            .expand(
                |index| {
                    if index == reference.level {
                        counter
                    } else {
                        ancestor_values[usize::from(index)]
                            .expect("referenced ancestor was validated")
                    }
                },
                |index| formats[usize::from(index)],
            )
            .map_err(counter_error)?;
        self.started.insert((reference.index, reference.level));
        if numberless {
            self.ancestors.insert(
                (reference.index, reference.level),
                AncestorValue {
                    value: None,
                    format: None,
                },
            );
        } else {
            self.ancestors.insert(
                (reference.index, reference.level),
                AncestorValue {
                    value: Some(counter),
                    format: Some(number_format::name(effective_format(
                        current, current, false,
                    ))?),
                },
            );
        }
        self.lists.insert(selection.list.id, state);
        Ok(NumberingInfo {
            num_id: u32::try_from(reference.index + 1)
                .map_err(|_| unsupported("Word list instance identifier overflow"))?,
            level: u32::from(reference.level),
            format: number_format::name(effective_format(current, current, false))?.to_string(),
            text,
            indent_left: paragraph.indent_left,
            tab: paragraph.indent_first.abs(),
            suff: ["tab", "space", "nothing"][usize::from(current.follow)].to_string(),
            jc: ["left", "center", "right"][usize::from(current.justification)].to_string(),
            font_family: facts.font_family.clone(),
            font_family_east_asia: facts.font_family_east_asia.clone(),
            font_facts: Some(facts),
            color: marker.direct_color(),
            color_auto: marker.direct_color_auto(),
            pic_bullet_image_path: None,
            pic_bullet_mime_type: None,
            pic_bullet_width_pt: None,
            pic_bullet_height_pt: None,
        })
    }
}

fn effective_format(current: &Level<'_>, referenced: &Level<'_>, inherited: bool) -> u8 {
    if current.legal && inherited && referenced.start.is_some() && referenced.format != 0x16 {
        0
    } else {
        referenced.format
    }
}

fn counter_error(error: ooxml_common::numbering::CounterError) -> String {
    match error {
        ooxml_common::numbering::CounterError::InvalidLevel => {
            unsupported("invalid Word list level")
        }
        ooxml_common::numbering::CounterError::Overflow => {
            unsupported("Word list counter overflow")
        }
        ooxml_common::numbering::CounterError::OutputTooLarge => {
            unsupported("Word list marker output too large")
        }
        ooxml_common::numbering::CounterError::InvalidTemplate => {
            unsupported("invalid Word numbering template")
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::numbering::{LevelOverride, List, Override};
    use std::collections::HashSet;

    const DECIMAL: &[u8] = &[0, 0, b'.', 0];
    const CHILD: &[u8] = &[0, 0, b'.', 0, 1, 0, b'.', 0];
    const LEVEL_ONE_ONLY: &[u8] = &[1, 0, b'.', 0];
    const LEVEL_TWO_ONLY: &[u8] = &[2, 0, b'.', 0];
    const BULLET: &[u8] = &[b'*', 0];

    fn level(index: u8, format: u8, start: Option<u16>, text: &'static [u8]) -> Level<'static> {
        let mut placeholders = [None; 9];
        if format != 0x17 && format != 0xff {
            placeholders[usize::from(index)] = Some((1 + index * 2, index));
            if index == 1 {
                placeholders[0] = Some((1, 0));
            } else if index == 2 {
                placeholders[0] = Some((1, 0));
                placeholders[1] = Some((3, 1));
            }
        }
        Level {
            start,
            format,
            justification: 0,
            legal: false,
            restart: (!matches!(format, 0x17 | 0xff)).then_some(index),
            follow: 0,
            tentative: false,
            papx: &[],
            chpx: &[],
            text,
            placeholders,
        }
    }

    fn own_counter_level(index: u8, start: u16) -> Level<'static> {
        let text = match index {
            1 => LEVEL_ONE_ONLY,
            2 => LEVEL_TWO_ONLY,
            _ => unreachable!(),
        };
        let mut value = level(index, 0, Some(start), text);
        value.placeholders = [None; 9];
        value.placeholders[usize::from(index)] = Some((1, index));
        value
    }

    fn tables(levels: Vec<Level<'static>>) -> Tables<'static> {
        Tables {
            lists: vec![List {
                id: 42,
                styles: [0xfff; 9],
                simple: levels.len() == 1,
                hybrid: levels.len() > 1,
                auto_number: false,
                levels,
            }],
            overrides: vec![Override {
                list_index: 0,
                first_cp: None,
                auto_number_field: None,
                levels: Vec::new(),
            }],
        }
    }

    fn activate(store: &mut Store, tables: &Tables<'_>, level: u8) -> String {
        activate_at(store, tables, 0, level)
    }

    fn activate_at(store: &mut Store, tables: &Tables<'_>, index: usize, level: u8) -> String {
        store
            .activate(
                tables,
                Reference {
                    index,
                    level,
                    preserve_indent: false,
                },
                &character::Properties::default(),
                &DocParagraph::default(),
                &[],
            )
            .unwrap()
            .text
    }

    #[test]
    fn sequences_once_per_activation_and_resets_at_story_boundary() {
        let tables = tables(vec![level(0, 0, Some(1), DECIMAL)]);
        let mut store = Store::default();
        store.begin_story().unwrap();
        assert_eq!(activate(&mut store, &tables, 0), "1.");
        assert_eq!(activate(&mut store, &tables, 0), "2.");
        store.begin_story().unwrap();
        assert_eq!(activate(&mut store, &tables, 0), "1.");
    }

    #[test]
    fn authoritative_offsets_do_not_reinterpret_literal_percent_text() {
        static TEXT: &[u8] = &[b'%', 0, b'1', 0, b' ', 0, 0, 0];
        let mut value = level(0, 0, Some(1), TEXT);
        value.placeholders[0] = Some((4, 0));
        let tables = tables(vec![value]);
        let mut store = Store::default();
        store.begin_story().unwrap();
        assert_eq!(activate(&mut store, &tables, 0), "%1 1");
    }

    #[test]
    fn legal_coerces_roman_ancestor_but_preserves_arabic_lz() {
        for (ancestor_format, expected, expected_format) in
            [(2, "1.1.", "decimal"), (0x16, "01.1.", "decimal")]
        {
            let mut levels = vec![
                level(0, ancestor_format, Some(1), DECIMAL),
                level(1, 0, Some(1), CHILD),
            ];
            levels[1].legal = true;
            let tables = tables(levels);
            let mut store = Store::default();
            store.begin_story().unwrap();
            activate(&mut store, &tables, 0);
            let info = store
                .activate(
                    &tables,
                    Reference {
                        index: 0,
                        level: 1,
                        preserve_indent: false,
                    },
                    &character::Properties::default(),
                    &DocParagraph::default(),
                    &[],
                )
                .unwrap();
            assert_eq!(info.text, expected);
            assert_eq!(info.format, expected_format);
        }
    }

    #[test]
    fn legal_current_arabic_lz_keeps_its_zero_padded_format() {
        let mut current = level(0, 0x16, Some(1), DECIMAL);
        current.legal = true;
        let tables = tables(vec![current]);
        let mut store = Store::default();
        store.begin_story().unwrap();
        let info = store
            .activate(
                &tables,
                Reference {
                    index: 0,
                    level: 0,
                    preserve_indent: false,
                },
                &character::Properties::default(),
                &DocParagraph::default(),
                &[],
            )
            .unwrap();
        assert_eq!(
            (info.format.as_str(), info.text.as_str()),
            ("decimalZero", "01.")
        );
    }

    #[test]
    fn legal_non_arabic_current_format_is_rejected_at_unmeasured_boundary() {
        let mut levels = vec![
            level(0, 0x16, Some(1), DECIMAL),
            level(1, 1, Some(1), CHILD),
        ];
        levels[1].legal = true;
        let tables = tables(levels);
        let mut store = Store::default();
        store.begin_story().unwrap();
        assert_eq!(activate(&mut store, &tables, 0), "01.");
        let error = store
            .activate(
                &tables,
                Reference {
                    index: 0,
                    level: 1,
                    preserve_indent: false,
                },
                &character::Properties::default(),
                &DocParagraph::default(),
                &[],
            )
            .unwrap_err();
        assert!(error.contains("legal Word list level with a non-Arabic current format"));
        // The rejected current level does not advance the shared sequence.
        assert_eq!(activate(&mut store, &tables, 0), "02.");
    }

    #[test]
    fn none_level_placeholder_is_validated_without_counter_callback() {
        static INVALID: &[u8] = &[1, 0];
        let mut none = level(0, 0xff, None, INVALID);
        none.placeholders[0] = Some((1, 1));
        let tables = tables(vec![none]);
        let mut store = Store::default();
        store.begin_story().unwrap();
        let error = store
            .activate(
                &tables,
                Reference {
                    index: 0,
                    level: 0,
                    preserve_indent: false,
                },
                &character::Properties::default(),
                &DocParagraph::default(),
                &[],
            )
            .unwrap_err();
        assert!(error.contains("placeholder offsets"));
    }

    #[test]
    fn bullet_parent_encounter_restarts_deeper_number_sequence() {
        let tables = tables(vec![
            level(0, 0x17, None, BULLET),
            level(1, 0, Some(1), CHILD),
        ]);
        let mut store = Store::default();
        store.begin_story().unwrap();
        assert_eq!(activate(&mut store, &tables, 0), "*");
        assert_eq!(activate(&mut store, &tables, 1), ".1.");
        assert_eq!(activate(&mut store, &tables, 1), ".2.");
        assert_eq!(activate(&mut store, &tables, 0), "*");
        assert_eq!(activate(&mut store, &tables, 1), ".1.");
    }

    #[test]
    fn same_level_numberless_paragraphs_count_but_unmeasured_override_is_rejected() {
        let mut tables = tables(vec![level(0, 0, Some(1), DECIMAL)]);
        tables.overrides.push(Override {
            list_index: 0,
            first_cp: None,
            auto_number_field: None,
            levels: vec![LevelOverride {
                index: 0,
                start: None,
                formatting: Some(level(0, 0x17, None, BULLET)),
            }],
        });
        tables.overrides.push(Override {
            list_index: 0,
            first_cp: None,
            auto_number_field: None,
            levels: vec![LevelOverride {
                index: 0,
                start: None,
                formatting: Some(level(0, 0xff, None, &[])),
            }],
        });
        let mut store = Store::default();
        store.begin_story().unwrap();
        assert_eq!(activate_at(&mut store, &tables, 0, 0), "1.");
        assert_eq!(activate_at(&mut store, &tables, 1, 0), "*");
        assert_eq!(activate_at(&mut store, &tables, 0, 0), "3.");
        assert_eq!(activate_at(&mut store, &tables, 2, 0), "");
        assert_eq!(activate_at(&mut store, &tables, 0, 0), "5.");

        tables.overrides[1].levels[0].start = Some(7);
        let mut store = Store::default();
        store.begin_story().unwrap();
        assert_eq!(activate_at(&mut store, &tables, 0, 0), "1.");
        let error = store
            .activate(
                &tables,
                Reference {
                    index: 1,
                    level: 0,
                    preserve_indent: false,
                },
                &character::Properties::default(),
                &DocParagraph::default(),
                &[],
            )
            .unwrap_err();
        assert!(error.contains("numberless Word list level with a start override"));
        // The unsupported activation is atomic and does not advance the shared
        // LSID sequence.
        assert_eq!(activate_at(&mut store, &tables, 0, 0), "2.");
    }

    #[test]
    fn bullet_model_preserves_unicode_and_symbol_font_glyph_references() {
        for (text, expected) in [(&[0x22, 0x20][..], "•"), (&[0x22, 0xf0][..], "\u{f022}")] {
            let tables = tables(vec![level(0, 0x17, None, text)]);
            let mut store = Store::default();
            store.begin_story().unwrap();
            let mut marker = character::Properties::default();
            marker.fonts[0] = Some(0);
            let info = store
                .activate(
                    &tables,
                    Reference {
                        index: 0,
                        level: 0,
                        preserve_indent: false,
                    },
                    &marker,
                    &DocParagraph::default(),
                    &["Symbol".to_string()],
                )
                .unwrap();
            assert_eq!(info.text, expected);
            assert_eq!(info.font_family.as_deref(), Some("Symbol"));
        }
    }

    #[test]
    fn many_lfo_instances_are_bounded_to_one_story_and_dropped_on_reset() {
        const COUNT: usize = 4096;
        let mut tables = tables(vec![level(0, 0, Some(1), DECIMAL)]);
        let prototype = || Override {
            list_index: 0,
            first_cp: None,
            auto_number_field: None,
            levels: Vec::new(),
        };
        tables.overrides.clear();
        tables.overrides.extend((0..COUNT).map(|_| prototype()));
        let mut store = Store::default();
        store.begin_story().unwrap();
        for index in 0..COUNT {
            store
                .activate(
                    &tables,
                    Reference {
                        index,
                        level: 0,
                        preserve_indent: false,
                    },
                    &character::Properties::default(),
                    &DocParagraph::default(),
                    &[],
                )
                .unwrap();
        }
        store.begin_story().unwrap();
        assert_eq!(activate(&mut store, &tables, 0), "1.");
    }

    #[test]
    fn full_level_and_start_override_are_selected_without_mutating_definition() {
        let mut tables = tables(vec![level(0, 0, Some(1), DECIMAL)]);
        tables.overrides[0].levels.push(LevelOverride {
            index: 0,
            start: Some(7),
            formatting: Some(level(0, 3, Some(1), DECIMAL)),
        });
        let mut store = Store::default();
        store.begin_story().unwrap();
        let info = store
            .activate(
                &tables,
                Reference {
                    index: 0,
                    level: 0,
                    preserve_indent: false,
                },
                &character::Properties::default(),
                &DocParagraph::default(),
                &[],
            )
            .unwrap();
        assert_eq!(
            (info.format.as_str(), info.text.as_str()),
            ("upperLetter", "G.")
        );
    }

    #[test]
    fn lfo_aliases_share_lsid_sequence_but_each_override_start_applies_once() {
        let mut tables = tables(vec![level(0, 0, Some(1), DECIMAL)]);
        tables.overrides.push(Override {
            list_index: 0,
            first_cp: None,
            auto_number_field: None,
            levels: vec![LevelOverride {
                index: 0,
                start: Some(5),
                formatting: None,
            }],
        });
        let mut store = Store::default();
        store.begin_story().unwrap();
        let mut marker = |index| {
            store
                .activate(
                    &tables,
                    Reference {
                        index,
                        level: 0,
                        preserve_indent: false,
                    },
                    &character::Properties::default(),
                    &DocParagraph::default(),
                    &[],
                )
                .unwrap()
                .text
        };
        assert_eq!(marker(0), "1.");
        assert_eq!(marker(1), "5.");
        assert_eq!(marker(0), "6.");
        assert_eq!(marker(1), "7.");
    }

    #[test]
    fn ancestor_placeholder_uses_same_lfo_value_and_its_own_format() {
        let tables = {
            let mut tables = tables(vec![
                level(0, 1, Some(1), DECIMAL),
                level(1, 0, Some(1), CHILD),
            ]);
            tables.overrides.push(Override {
                list_index: 0,
                first_cp: None,
                auto_number_field: None,
                levels: vec![LevelOverride {
                    index: 0,
                    start: None,
                    formatting: Some(level(0, 0, Some(1), DECIMAL)),
                }],
            });
            tables
        };
        let mut store = Store::default();
        store.begin_story().unwrap();
        assert_eq!(activate_at(&mut store, &tables, 0, 0), "I.");
        // The numeric sequence is shared by LSID, so the alias is 2.
        assert_eq!(activate_at(&mut store, &tables, 1, 0), "2.");
        // The ancestor placeholder belongs to iLfo 0, so it uses that list's
        // closest prior value (1) and upper-Roman format, not alias value 2.
        assert_eq!(activate_at(&mut store, &tables, 0, 1), "I.1.");
    }

    #[test]
    fn absent_ancestor_is_not_fabricated_as_one() {
        let tables = tables(vec![
            level(0, 0, Some(1), DECIMAL),
            level(1, 0, Some(1), CHILD),
        ]);
        let mut store = Store::default();
        store.begin_story().unwrap();
        let error = store
            .activate(
                &tables,
                Reference {
                    index: 0,
                    level: 1,
                    preserve_indent: false,
                },
                &character::Properties::default(),
                &DocParagraph::default(),
                &[],
            )
            .unwrap_err();
        assert!(error.contains("missing prior Word list ancestor level"));
    }

    #[test]
    fn current_lfo_effective_restart_threshold_controls_shared_lsid_descendants() {
        // Literal bounded reference for MS-DOC 2.4.6.4 steps 1-13: find the
        // closest prior paragraph at this level, inspect intervening real
        // paragraph levels against the CURRENT LVL's restart limit, then
        // increment or restart. These fixtures use current-level-only marker
        // templates so the counter oracle stays independent of ancestor lookup.
        #[derive(Default)]
        struct ReferenceAlgorithm {
            history: Vec<(u8, Option<u32>)>,
        }
        impl ReferenceAlgorithm {
            fn advance(
                &mut self,
                level: u8,
                restart: u8,
                start: u32,
                absolute: Option<u32>,
            ) -> u32 {
                let mut relative_count = 0u32;
                let mut absolute_value = None;
                for &(prior_level, prior_absolute) in &self.history {
                    if prior_level < restart {
                        relative_count = 0;
                        absolute_value = None;
                    } else if prior_level == level {
                        if let Some(override_value) = prior_absolute {
                            absolute_value = Some(override_value);
                            relative_count = 0;
                        } else if let Some(value) = absolute_value {
                            absolute_value = Some(value + 1);
                        } else {
                            relative_count += 1;
                        }
                    }
                }
                let value = absolute.unwrap_or_else(|| {
                    absolute_value.map_or(start + relative_count, |value| value + 1)
                });
                self.history.push((level, absolute));
                value
            }
            fn text(&self, value: u32) -> String {
                format!("{value}.")
            }
        }
        let mut tables = tables(vec![
            level(0, 0, Some(1), DECIMAL),
            own_counter_level(1, 1),
            own_counter_level(2, 1),
        ]);
        let mut no_level_one_restart = own_counter_level(2, 1);
        no_level_one_restart.restart = Some(1);
        tables.overrides[0].levels.push(LevelOverride {
            index: 2,
            start: None,
            formatting: Some(no_level_one_restart),
        });
        let mut level_one_restarts = own_counter_level(2, 1);
        level_one_restarts.restart = Some(2);
        tables.overrides.push(Override {
            list_index: 0,
            first_cp: None,
            auto_number_field: None,
            levels: vec![LevelOverride {
                index: 2,
                start: None,
                formatting: Some(level_one_restarts),
            }],
        });
        let mut store = Store::default();
        let mut oracle = ReferenceAlgorithm::default();
        store.begin_story().unwrap();
        for (instance, level, restart) in [(0, 2, 1), (1, 1, 1), (0, 2, 1)] {
            let value = oracle.advance(level, restart, 1, None);
            assert_eq!(
                activate_at(&mut store, &tables, instance, level),
                oracle.text(value)
            );
        }

        store.begin_story().unwrap();
        let mut oracle = ReferenceAlgorithm::default();
        for (instance, level, restart) in [(1, 2, 2), (0, 1, 1), (1, 2, 2)] {
            let value = oracle.advance(level, restart, 1, None);
            assert_eq!(
                activate_at(&mut store, &tables, instance, level),
                oracle.text(value)
            );
        }

        store.begin_story().unwrap();
        let mut oracle = ReferenceAlgorithm::default();
        for (instance, level, restart) in [(0, 2, 1), (0, 1, 1), (1, 2, 2), (0, 2, 1)] {
            let value = oracle.advance(level, restart, 1, None);
            assert_eq!(
                activate_at(&mut store, &tables, instance, level),
                oracle.text(value)
            );
        }
    }

    #[test]
    fn current_full_level_start_is_the_base_for_prior_relative_occurrences() {
        let mut tables = tables(vec![level(0, 0, Some(1), DECIMAL)]);
        tables.overrides[0].levels.push(LevelOverride {
            index: 0,
            start: None,
            formatting: Some(level(0, 0, Some(1), DECIMAL)),
        });
        tables.overrides.push(Override {
            list_index: 0,
            first_cp: None,
            auto_number_field: None,
            levels: vec![LevelOverride {
                index: 0,
                start: None,
                formatting: Some(level(0, 0, Some(3), DECIMAL)),
            }],
        });
        let mut store = Store::default();
        store.begin_story().unwrap();
        assert_eq!(activate_at(&mut store, &tables, 0, 0), "1.");
        assert_eq!(activate_at(&mut store, &tables, 1, 0), "4.");
        assert_eq!(activate_at(&mut store, &tables, 0, 0), "3.");
    }

    #[test]
    fn unsupported_auto_fields_and_formats_fail_without_partial_counter_commit() {
        let mut tables = tables(vec![level(0, 60, Some(1), DECIMAL)]);
        let mut store = Store::default();
        store.begin_story().unwrap();
        assert!(store
            .activate(
                &tables,
                Reference {
                    index: 0,
                    level: 0,
                    preserve_indent: false
                },
                &character::Properties::default(),
                &DocParagraph::default(),
                &[],
            )
            .unwrap_err()
            .contains("number format"));
        tables.lists[0].levels[0].format = 0;
        tables.overrides[0].auto_number_field = Some(0xfc);
        assert!(store
            .activate(
                &tables,
                Reference {
                    index: 0,
                    level: 0,
                    preserve_indent: false
                },
                &character::Properties::default(),
                &DocParagraph::default(),
                &[],
            )
            .unwrap_err()
            .contains("automatic-number"));
        tables.overrides[0].auto_number_field = None;
        assert_eq!(activate(&mut store, &tables, 0), "1.");
    }

    #[test]
    fn restart_start_and_once_only_override_summaries_match_exhaustive_replay() {
        #[derive(Clone, Copy)]
        struct Event {
            instance: usize,
            level: u8,
            restart: u8,
            start: u32,
            start_override: Option<u32>,
        }
        let events = [
            Event {
                instance: 0,
                level: 1,
                restart: 1,
                start: 1,
                start_override: None,
            },
            Event {
                instance: 0,
                level: 2,
                restart: 1,
                start: 1,
                start_override: None,
            },
            Event {
                instance: 1,
                level: 2,
                restart: 2,
                start: 3,
                start_override: Some(7),
            },
        ];
        let mut tables = tables(vec![
            level(0, 0, Some(1), DECIMAL),
            own_counter_level(1, 1),
            own_counter_level(2, 1),
        ]);
        let mut primary_level = own_counter_level(2, 1);
        primary_level.restart = Some(1);
        tables.overrides[0].levels.push(LevelOverride {
            index: 2,
            start: None,
            formatting: Some(primary_level),
        });
        let mut alias_level = own_counter_level(2, 3);
        alias_level.restart = Some(2);
        tables.overrides.push(Override {
            list_index: 0,
            first_cp: None,
            auto_number_field: None,
            levels: vec![LevelOverride {
                index: 2,
                start: Some(7),
                formatting: Some(alias_level),
            }],
        });

        for mut selector in 0..3usize.pow(6) {
            let mut store = Store::default();
            store.begin_story().unwrap();
            let mut history: Vec<(u8, Option<u32>)> = Vec::new();
            let mut started = HashSet::new();
            for _ in 0..6 {
                let event = events[selector % events.len()];
                selector /= events.len();
                let first = started.insert((event.instance, event.level));
                let absolute = first.then_some(event.start_override).flatten();
                let mut relative_count = 0u32;
                let mut absolute_value = None;
                for &(prior_level, prior_absolute) in &history {
                    if prior_level < event.restart {
                        relative_count = 0;
                        absolute_value = None;
                    } else if prior_level == event.level {
                        if let Some(override_value) = prior_absolute {
                            absolute_value = Some(override_value);
                            relative_count = 0;
                        } else if let Some(value) = absolute_value {
                            absolute_value = Some(value + 1);
                        } else {
                            relative_count += 1;
                        }
                    }
                }
                let value = absolute.unwrap_or_else(|| {
                    absolute_value.map_or(event.start + relative_count, |value| value + 1)
                });
                history.push((event.level, absolute));
                let expected = format!("{value}.");
                assert_eq!(
                    activate_at(&mut store, &tables, event.instance, event.level),
                    expected
                );
            }
        }
    }
}
