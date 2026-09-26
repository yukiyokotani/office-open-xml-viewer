//! Internal, format-agnostic numbering counter state.
//!
//! Acquisition adapters retain their own level/style structures and lend only
//! the effective counter facts to this engine. This keeps XML parsing and DOC
//! list-table parsing out of the shared layer.

use std::collections::{HashMap, HashSet};
use std::hash::Hash;

mod format;
#[doc(hidden)]
pub use format::format_counter;
#[doc(hidden)]
pub use format::format_word_synthetic_zero;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CounterError {
    Overflow,
    OutputTooLarge,
    InvalidTemplate,
}

/// Per-marker retained UTF-8 ceiling. This is an implementation resource
/// policy, not an OOXML schema limit or compatibility heuristic.
#[doc(hidden)]
pub const MAX_MARKER_BYTES: usize = 64 * 1024;

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum TemplatePart {
    Literal(String),
    Counter(u8),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NumberingTemplate {
    parts: Vec<TemplatePart>,
}

impl NumberingTemplate {
    /// Resource policy for the typed acquisition seam: at most nine explicit
    /// references and their ten intervening literal spans. The referenced level
    /// numbers are the source adapter's to validate. This does not constrain
    /// repeated `%N` occurrences in the separate OOXML text adapter.
    pub fn new(parts: Vec<TemplatePart>) -> Result<Self, CounterError> {
        if parts.len() > 19
            || parts
                .iter()
                .filter(|part| matches!(part, TemplatePart::Counter(_)))
                .count()
                > 9
        {
            return Err(CounterError::InvalidTemplate);
        }
        let template = Self { parts };
        if template.literal_bytes()? > MAX_MARKER_BYTES {
            return Err(CounterError::OutputTooLarge);
        }
        Ok(template)
    }

    fn literal_bytes(&self) -> Result<usize, CounterError> {
        self.parts
            .iter()
            .try_fold(0usize, |total, part| match part {
                TemplatePart::Literal(value) => total
                    .checked_add(value.len())
                    .ok_or(CounterError::OutputTooLarge),
                TemplatePart::Counter(_) => Ok(total),
            })
    }

    /// Expand explicit counter parts. The adapter supplies the effective format
    /// for each reference, including any source-format-specific legal policy.
    pub fn expand<'a>(
        &self,
        mut value_at: impl FnMut(u8) -> u32,
        mut format_at: impl FnMut(u8) -> &'a str,
    ) -> Result<String, CounterError> {
        let mut output = String::new();
        for part in &self.parts {
            match part {
                TemplatePart::Literal(value) => {
                    if value.len() > MAX_MARKER_BYTES.saturating_sub(output.len()) {
                        return Err(CounterError::OutputTooLarge);
                    }
                    output.push_str(value);
                }
                TemplatePart::Counter(level) => {
                    let value = format::format_counter_bounded(
                        value_at(*level),
                        format_at(*level),
                        MAX_MARKER_BYTES,
                    )
                    .map_err(|()| CounterError::OutputTooLarge)?;
                    if value.len() > MAX_MARKER_BYTES.saturating_sub(output.len()) {
                        return Err(CounterError::OutputTooLarge);
                    }
                    output.push_str(&value);
                }
            }
        }
        Ok(output)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum CounterIdentity<Shared, Orphan> {
    Shared(Shared),
    Orphan(Orphan),
}

#[derive(Clone, Copy, Debug)]
pub struct LevelFacts<'a> {
    pub start: u32,
    pub restart: Option<u32>,
    pub format: &'a str,
    pub text: &'a str,
    pub legal: bool,
}

#[derive(Clone, Debug)]
pub struct CounterEngine<Shared, Instance, Orphan> {
    counters: HashMap<CounterIdentity<Shared, Orphan>, LevelCounters>,
    started: HashSet<(Instance, u32)>,
}

/// Live counters of one counter identity. A level is live when it has an
/// explicit value or lies below `seeded_below`; an implicitly live level
/// shows 1.
///
/// Advancing level L seeds every shallower level to its start. Materialising
/// that range would make the work and memory of one paragraph proportional to
/// its authored level number, so a huge level would exhaust the WASM heap.
/// Only the caller's `specific_levels` (levels with their own definition or
/// start override) are stored; every other seeded level starts at 1 (the
/// default start) and is represented by the range. The observable counters are
/// unchanged. How Word displays a level beyond its nine defined levels is not
/// modelled here.
#[derive(Clone, Debug, Default)]
struct LevelCounters {
    values: HashMap<u32, u32>,
    seeded_below: u32,
}

impl LevelCounters {
    fn get(&self, level: u32) -> Option<u32> {
        match self.values.get(&level) {
            Some(&value) => Some(value),
            None if level < self.seeded_below => Some(1),
            None => None,
        }
    }
}

impl<Shared, Instance, Orphan> Default for CounterEngine<Shared, Instance, Orphan> {
    fn default() -> Self {
        Self {
            counters: HashMap::new(),
            started: HashSet::new(),
        }
    }
}

impl<Shared, Instance, Orphan> CounterEngine<Shared, Instance, Orphan>
where
    Shared: Copy + Eq + Hash,
    Instance: Copy + Eq + Hash,
    Orphan: Copy + Eq + Hash,
{
    /// Advance one effective level and return the value to display.
    ///
    /// `specific_levels` lists the levels whose start or restart can differ
    /// from the defaults for this instance (own definition or start override).
    /// Work is proportional to those levels and to the live counters, never to
    /// the authored level number (see [`LevelCounters`]). A start override
    /// applies on the instance's first appearance at the level. OOXML's
    /// counter domain is broader than this model's u32, so a counter saturates
    /// at the representable boundary rather than trapping on a hostile document
    /// with a maximal startOverride.
    #[allow(clippy::too_many_arguments)]
    pub fn advance<'a>(
        &mut self,
        identity: CounterIdentity<Shared, Orphan>,
        instance: Instance,
        level: u32,
        start_override: Option<u32>,
        specific_levels: &[u32],
        mut start_at: impl FnMut(u32) -> u32,
        mut effective_level: impl FnMut(u32) -> Option<LevelFacts<'a>>,
    ) -> u32 {
        // `insert` returns true when the pair was NOT already present.
        let first_for_instance = self.started.insert((instance, level));
        let mut threshold = |deeper: u32| {
            effective_level(deeper)
                .and_then(|facts| facts.restart)
                .filter(|&value| value <= deeper)
                .unwrap_or(deeper)
        };

        // Resolve each live descendant's own policy, including a complete
        // level replacement. A never-restarting parent does not shield its
        // children, and merely seeding an ancestor is not an occurrence of it.
        // Iterate existing counters, never an input-provided restart range.
        let empty = LevelCounters::default();
        let current = self.counters.get(&identity).unwrap_or(&empty);
        let resets: Vec<u32> = current
            .values
            .keys()
            .copied()
            .filter(|&deeper| deeper > level && level < threshold(deeper))
            .collect();
        // Implicitly live levels deeper than `level` are reset too, except a
        // level whose own definition keeps it; such a level stays live at 1.
        let kept: Vec<u32> = specific_levels
            .iter()
            .copied()
            .filter(|&deeper| {
                deeper > level
                    && deeper < current.seeded_below
                    && !current.values.contains_key(&deeper)
                    && level >= threshold(deeper)
            })
            .collect();
        // Seed shallower levels to their start (their displayed value when they
        // are never advanced themselves), but never clobber a live ancestor.
        let seeded_below = match level.checked_add(1) {
            Some(next) => current.seeded_below.min(next),
            None => current.seeded_below,
        };
        let seeds: Vec<(u32, u32)> = specific_levels
            .iter()
            .copied()
            .filter(|&l| l >= seeded_below && l < level && !current.values.contains_key(&l))
            .map(|l| (l, start_at(l)))
            .collect();

        let value = match start_override.filter(|_| first_for_instance) {
            Some(start) => start,
            None => match current
                .values
                .get(&level)
                .copied()
                .or_else(|| (level < seeded_below).then_some(1))
            {
                Some(value) => value.saturating_add(1),
                None => start_at(level),
            },
        };

        let entry = self.counters.entry(identity).or_default();
        for deeper in resets {
            entry.values.remove(&deeper);
        }
        for deeper in kept {
            entry.values.insert(deeper, 1);
        }
        for (ancestor, start) in seeds {
            entry.values.insert(ancestor, start);
        }
        entry.seeded_below = seeded_below.max(level);
        entry.values.insert(level, value);
        value
    }

    /// Resolve `%1`…`%9` from deepest to shallowest so placeholders cannot
    /// partially overlap, using the shared ST_NumberFormat formatter. Only
    /// placeholders present in the text are visited, so the work does not grow
    /// with the authored level number.
    pub fn resolve_text<'a>(
        &self,
        identity: CounterIdentity<Shared, Orphan>,
        level: u32,
        counter: u32,
        start_at: impl FnMut(u32) -> u32,
        effective_level: impl FnMut(u32) -> Option<LevelFacts<'a>>,
    ) -> String {
        self.resolve_text_with_zero_mode(identity, level, counter, false, start_at, effective_level)
    }

    /// Resolve a marker whose own level shows Word's synthetic zero (see
    /// [`word_level_use`]). This is a display-only value, not a live counter;
    /// a reset placeholder also shows its format's synthetic zero until its
    /// counter becomes live.
    pub fn resolve_text_word_zero<'a>(
        &self,
        identity: CounterIdentity<Shared, Orphan>,
        level: u32,
        start_at: impl FnMut(u32) -> u32,
        effective_level: impl FnMut(u32) -> Option<LevelFacts<'a>>,
    ) -> String {
        self.resolve_text_with_zero_mode(identity, level, 0, true, start_at, effective_level)
    }

    fn resolve_text_with_zero_mode<'a>(
        &self,
        identity: CounterIdentity<Shared, Orphan>,
        level: u32,
        counter: u32,
        word_zero: bool,
        mut start_at: impl FnMut(u32) -> u32,
        mut effective_level: impl FnMut(u32) -> Option<LevelFacts<'a>>,
    ) -> String {
        let Some(current) = effective_level(level) else {
            return format!("{counter}.");
        };
        let mut text = current.text.to_owned();
        let mut highest = Some(level);
        while let Some(ancestor) = highest.and_then(|hi| deepest_placeholder(&text, hi)) {
            let live = if ancestor == level {
                Some(counter)
            } else {
                self.counters
                    .get(&identity)
                    .and_then(|counts| counts.get(ancestor))
            };
            let value = live.unwrap_or_else(|| start_at(ancestor));
            // 17.9.4 applies to this marker's entire displayed level text,
            // including its own placeholder. Keep authored formats intact so
            // other markers continue to use their own definitions.
            let format = if current.legal {
                "decimal"
            } else {
                effective_level(ancestor)
                    .map(|facts| facts.format)
                    .unwrap_or(current.format)
            };
            let rendered = if format == "bullet" {
                // A bullet uses literal lvlText; Word suppresses a numeric
                // placeholder even when the level text contains one.
                String::new()
            } else if word_zero && (ancestor == level || live.is_none()) {
                format::format_word_synthetic_zero(format)
            } else {
                format::format_counter(value, format)
            };
            text = text.replace(&format!("%{}", u64::from(ancestor) + 1), &rendered);
            highest = ancestor.checked_sub(1);
        }
        text
    }

    /// Expand an explicitly segmented source template against this engine's
    /// live counters. Placeholder identity comes from source offsets, never
    /// from scanning literal text.
    pub fn resolve_template<'a>(
        &self,
        identity: CounterIdentity<Shared, Orphan>,
        level: u32,
        counter: u32,
        template: &NumberingTemplate,
        mut start_at: impl FnMut(u8) -> u32,
        format_at: impl FnMut(u8) -> &'a str,
    ) -> Result<String, CounterError> {
        if template.parts.iter().any(
            |part| matches!(part, TemplatePart::Counter(referenced) if u32::from(*referenced) > level),
        ) {
            return Err(CounterError::InvalidTemplate);
        }
        template.expand(
            |referenced| {
                if u32::from(referenced) == level {
                    counter
                } else {
                    self.counters
                        .get(&identity)
                        .and_then(|counts| counts.get(u32::from(referenced)))
                        .unwrap_or_else(|| start_at(referenced))
                }
            },
            format_at,
        )
    }
}

/// Word's paragraph level byte (DOCX `w:ilvl` after its lexical narrowing;
/// see the DOCX adapter) mapped to the level that paints the marker and the
/// level whose counter advances. Separate the formatting source from the
/// advancing counter. ECMA-376 Part 1
/// §17.9 defines only the reference/definition relationship; MS-OI29500
/// §2.1.277 documents Word's 0..255 authoring limit, not its malformed-value
/// rendering. These branches are observations from Word 16.113 controls:
/// every byte 0..255; signed, leading-zero, plus, 2^31 and 2^32 boundaries;
/// a two-level definition; a level-8 replacement; and style-origin numPr.
/// Each control placed levels 0, 1, and 8 before/after the target and checked
/// its marker and subsequent counters. Values 16..255 repeat every 16 bytes;
/// 9..15 have the special first-block behavior encoded below. The fixed `9`
/// is a fresh fallback marker value, even when level 8 has startOverride=20.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WordLevelUse {
    pub marker_level: Option<u32>,
    pub counter_level: Option<u32>,
    /// `Some(0)` or `Some(9)` uses a display value independent of any counter.
    pub fixed_counter: Option<u32>,
}

pub fn word_level_use(ilvl: u8) -> WordLevelUse {
    let (marker_level, counter_level, fixed_counter) = match ilvl {
        0..=8 => (Some(u32::from(ilvl)), Some(u32::from(ilvl)), None),
        9..=12 => (None, None, None),
        13..=14 => (Some(8), None, Some(9)),
        15 => (Some(0), Some(0), None),
        _ => match ilvl % 16 {
            level @ 0..=7 => (Some(8), Some(u32::from(level)), Some(0)),
            8 => (Some(8), Some(8), None),
            _ => (Some(8), None, Some(9)),
        },
    };
    WordLevelUse {
        marker_level,
        counter_level,
        fixed_counter,
    }
}

/// The deepest level `k <= highest` whose placeholder `%{k+1}` occurs in
/// `text`. Every `%` followed by digits contributes each digit prefix without a
/// leading zero, because `%1` also matches inside `%12`.
fn deepest_placeholder(text: &str, highest: u32) -> Option<u32> {
    let bytes = text.as_bytes();
    let mut best = None;
    for (at, _) in text.match_indices('%') {
        let mut number: u64 = 0;
        for &digit in bytes[at + 1..].iter().take(10) {
            if !digit.is_ascii_digit() || (number == 0 && digit == b'0') {
                break;
            }
            number = number * 10 + u64::from(digit - b'0');
            if let Ok(k) = u32::try_from(number - 1) {
                if k <= highest && best.is_none_or(|b| k > b) {
                    best = Some(k);
                }
            }
        }
    }
    best
}

#[cfg(test)]
mod tests {
    use super::*;

    fn decimal(start: u32, restart: Option<u32>) -> LevelFacts<'static> {
        LevelFacts {
            start,
            restart,
            format: "decimal",
            text: "%1.%2",
            legal: false,
        }
    }

    #[test]
    fn aliases_share_but_orphans_are_disjoint_and_override_starts_once() {
        let mut state = CounterEngine::<u32, u32, u32>::default();
        let levels = |level| (level == 0).then(|| decimal(1, None));
        let shared = CounterIdentity::Shared(10);
        assert_eq!(state.advance(shared, 1, 0, None, &[0], |_| 1, levels), 1);
        assert_eq!(state.advance(shared, 2, 0, Some(7), &[0], |_| 1, levels), 7);
        assert_eq!(state.advance(shared, 1, 0, None, &[0], |_| 1, levels), 8);
        assert_eq!(state.advance(shared, 2, 0, Some(7), &[0], |_| 1, levels), 9);
        assert_eq!(
            state.advance(CounterIdentity::Orphan(10), 10, 0, None, &[], |_| 1, levels),
            1
        );
    }

    #[test]
    fn ancestors_seed_and_restart_descendants() {
        let mut state = CounterEngine::<u32, u32, u32>::default();
        let facts = |level| match level {
            0 => Some(LevelFacts {
                text: "%1",
                ..decimal(3, None)
            }),
            1 => Some(LevelFacts {
                text: "%1.%2",
                ..decimal(1, Some(1))
            }),
            _ => None,
        };
        let id = CounterIdentity::Shared(4);
        let start = |level| facts(level).map_or(1, |f| f.start);
        let child = state.advance(id, 1, 1, None, &[0, 1], start, facts);
        assert_eq!(state.resolve_text(id, 1, child, start, facts), "3.1");
        state.advance(id, 1, 0, None, &[0, 1], start, facts);
        assert_eq!(state.advance(id, 1, 1, None, &[0, 1], start, facts), 1);
    }

    #[test]
    fn missing_definition_keeps_fallback_text_but_honors_override_start() {
        let mut state = CounterEngine::<u32, u32, u32>::default();
        let id = CounterIdentity::Orphan(9);
        let value = state.advance(
            id,
            9,
            1,
            Some(7),
            &[1],
            |level| if level == 0 { 3 } else { 7 },
            |_| None,
        );
        assert_eq!(value, 7);
        assert_eq!(state.resolve_text(id, 1, value, |_| 1, |_| None), "7.");
    }

    #[test]
    fn unseeded_ancestor_uses_instance_start_override() {
        let state = CounterEngine::<u32, u32, u32>::default();
        let id = CounterIdentity::Shared(2);
        let facts = |level| match level {
            0 => Some(LevelFacts {
                text: "%1",
                ..decimal(1, None)
            }),
            1 => Some(LevelFacts {
                text: "%1.%2",
                ..decimal(1, None)
            }),
            _ => None,
        };
        assert_eq!(
            state.resolve_text(id, 1, 4, |level| if level == 0 { 7 } else { 1 }, facts),
            "7.4"
        );
    }

    #[test]
    fn typed_template_preserves_literal_percent_digits_and_unicode() {
        let template = NumberingTemplate::new(vec![
            TemplatePart::Literal("A%1😀".to_string()),
            TemplatePart::Counter(0),
            TemplatePart::Literal("Z".to_string()),
        ])
        .unwrap();
        assert_eq!(template.expand(|_| 7, |_| "decimal").unwrap(), "A%1😀7Z");
        let engine = CounterEngine::<u32, u32, u32>::default();
        assert_eq!(
            engine
                .resolve_template(
                    CounterIdentity::Shared(1),
                    0,
                    7,
                    &template,
                    |_| 1,
                    |_| "decimal",
                )
                .unwrap(),
            "A%1😀7Z"
        );
    }

    #[test]
    fn typed_template_rejects_excessive_or_future_references() {
        assert_eq!(
            NumberingTemplate::new(vec![TemplatePart::Counter(0); 10]),
            Err(CounterError::InvalidTemplate)
        );
        let template = NumberingTemplate::new(vec![TemplatePart::Counter(1)]).unwrap();
        let engine = CounterEngine::<u32, u32, u32>::default();
        assert_eq!(
            engine.resolve_template(
                CounterIdentity::Shared(1),
                0,
                1,
                &template,
                |_| 1,
                |_| "decimal"
            ),
            Err(CounterError::InvalidTemplate)
        );
    }

    #[test]
    fn explicit_template_expansion_obeys_exact_output_budget() {
        let exact = "x".repeat(MAX_MARKER_BYTES);
        assert_eq!(
            NumberingTemplate::new(vec![TemplatePart::Literal(exact)])
                .unwrap()
                .expand(|_| 1, |_| "decimal")
                .unwrap()
                .len(),
            MAX_MARKER_BYTES
        );
        let over = "x".repeat(MAX_MARKER_BYTES + 1);
        assert_eq!(
            NumberingTemplate::new(vec![TemplatePart::Literal(over)]),
            Err(CounterError::OutputTooLarge)
        );
    }
}
