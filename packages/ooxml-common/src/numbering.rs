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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CounterError {
    /// Outside the engine's supported nine-level (0..=8) resource policy.
    /// This is not an XSD restriction on ST_DecimalNumber.
    InvalidLevel,
    Overflow,
    OutputTooLarge,
}

/// Per-marker retained UTF-8 ceiling. This is an implementation resource
/// policy, not an OOXML schema limit or compatibility heuristic.
const MAX_MARKER_BYTES: usize = 64 * 1024;

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
    counters: HashMap<CounterIdentity<Shared, Orphan>, HashMap<u32, u32>>,
    started: HashSet<(Instance, u32)>,
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
    /// Advance one effective level. The caller supplies a bounded lookup over
    /// already-acquired levels; this method scans only ancestors and live
    /// descendants already present in the counter map.
    pub fn advance<'a>(
        &mut self,
        identity: CounterIdentity<Shared, Orphan>,
        instance: Instance,
        level: u32,
        start_override: Option<u32>,
        mut start_at: impl FnMut(u32) -> u32,
        mut effective_level: impl FnMut(u32) -> Option<LevelFacts<'a>>,
    ) -> Result<u32, CounterError> {
        if level > 8 {
            return Err(CounterError::InvalidLevel);
        }
        let starts: Vec<u32> = (0..=level).map(&mut start_at).collect();
        let first_for_instance = !self.started.contains(&(instance, level));
        let resets: Vec<u32> = self
            .counters
            .get(&identity)
            .into_iter()
            .flat_map(|counts| counts.keys().copied())
            .filter(|&deeper| {
                let threshold = effective_level(deeper)
                    .and_then(|facts| facts.restart)
                    .filter(|&value| value <= deeper)
                    .unwrap_or(deeper);
                deeper > level && level < threshold
            })
            .collect();
        let value = if first_for_instance && start_override.is_some() {
            start_override.expect("checked above")
        } else {
            match self
                .counters
                .get(&identity)
                .and_then(|entry| entry.get(&level))
            {
                Some(value) => value.checked_add(1).ok_or(CounterError::Overflow)?,
                None => starts[level as usize],
            }
        };
        // Commit only after every fallible calculation succeeds.
        self.started.insert((instance, level));
        let entry = self.counters.entry(identity).or_default();
        for deeper in resets {
            entry.remove(&deeper);
        }
        for (ancestor, &start) in starts.iter().enumerate().take(level as usize) {
            entry.entry(ancestor as u32).or_insert(start);
        }
        entry.insert(level, value);
        Ok(value)
    }

    /// Resolve `%1`…`%9` from deepest to shallowest so placeholders cannot
    /// partially overlap, using the shared ST_NumberFormat formatter.
    pub fn resolve_text<'a>(
        &self,
        identity: CounterIdentity<Shared, Orphan>,
        level: u32,
        counter: u32,
        mut start_at: impl FnMut(u32) -> u32,
        mut effective_level: impl FnMut(u32) -> Option<LevelFacts<'a>>,
    ) -> Result<String, CounterError> {
        if level > 8 {
            return Err(CounterError::InvalidLevel);
        }
        let Some(current) = effective_level(level) else {
            return Ok(format!("{counter}."));
        };
        if current.text.len() > MAX_MARKER_BYTES {
            return Err(CounterError::OutputTooLarge);
        }
        let mut text = current.text.to_owned();
        for ancestor in (0..=level).rev() {
            let value = if ancestor == level {
                counter
            } else {
                self.counters
                    .get(&identity)
                    .and_then(|counts| counts.get(&ancestor))
                    .copied()
                    .unwrap_or_else(|| start_at(ancestor))
            };
            let format = if current.legal {
                "decimal"
            } else {
                effective_level(ancestor)
                    .map(|facts| facts.format)
                    .unwrap_or(current.format)
            };
            let placeholder = format!("%{}", ancestor + 1);
            let occurrences = text.match_indices(&placeholder).count();
            if occurrences == 0 {
                continue;
            }
            let replacement = format::format_counter_bounded(value, format, MAX_MARKER_BYTES)
                .map_err(|()| CounterError::OutputTooLarge)?;
            let removed = occurrences
                .checked_mul(placeholder.len())
                .ok_or(CounterError::OutputTooLarge)?;
            let inserted = occurrences
                .checked_mul(replacement.len())
                .ok_or(CounterError::OutputTooLarge)?;
            let resulting = text
                .len()
                .checked_sub(removed)
                .and_then(|len| len.checked_add(inserted))
                .filter(|len| *len <= MAX_MARKER_BYTES)
                .ok_or(CounterError::OutputTooLarge)?;
            debug_assert_eq!(resulting, text.len() - removed + inserted);
            text = text.replace(&placeholder, &replacement);
        }
        Ok(text)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::cell::Cell;

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
        assert_eq!(state.advance(shared, 1, 0, None, |_| 1, levels), Ok(1));
        assert_eq!(state.advance(shared, 2, 0, Some(7), |_| 1, levels), Ok(7));
        assert_eq!(state.advance(shared, 1, 0, None, |_| 1, levels), Ok(8));
        assert_eq!(state.advance(shared, 2, 0, Some(7), |_| 1, levels), Ok(9));
        assert_eq!(
            state.advance(CounterIdentity::Orphan(10), 10, 0, None, |_| 1, levels),
            Ok(1)
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
        let child = state
            .advance(
                id,
                1,
                1,
                None,
                |level| facts(level).map_or(1, |f| f.start),
                facts,
            )
            .unwrap();
        assert_eq!(
            state.resolve_text(
                id,
                1,
                child,
                |level| facts(level).map_or(1, |f| f.start),
                facts
            ),
            Ok("3.1".to_string())
        );
        state
            .advance(
                id,
                1,
                0,
                None,
                |level| facts(level).map_or(1, |f| f.start),
                facts,
            )
            .unwrap();
        assert_eq!(
            state.advance(
                id,
                1,
                1,
                None,
                |level| facts(level).map_or(1, |f| f.start),
                facts
            ),
            Ok(1)
        );
    }

    #[test]
    fn missing_definition_keeps_fallback_text_but_honors_override_start() {
        let mut state = CounterEngine::<u32, u32, u32>::default();
        let id = CounterIdentity::Orphan(9);
        let value = state
            .advance(
                id,
                9,
                1,
                Some(7),
                |level| if level == 0 { 3 } else { 7 },
                |_| None,
            )
            .unwrap();
        assert_eq!(value, 7);
        assert_eq!(
            state.resolve_text(id, 1, value, |_| 1, |_| None),
            Ok("7.".to_string())
        );
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
            Ok("7.4".to_string())
        );
    }

    #[test]
    fn invalid_levels_fail_before_callbacks_or_mutation() {
        let mut state = CounterEngine::<u32, u32, u32>::default();
        let called = Cell::new(false);
        let id = CounterIdentity::Shared(1);
        assert_eq!(
            state.advance(
                id,
                1,
                9,
                None,
                |_| {
                    called.set(true);
                    1
                },
                |_| {
                    called.set(true);
                    None
                }
            ),
            Err(CounterError::InvalidLevel)
        );
        assert_eq!(
            state.resolve_text(
                id,
                u32::MAX,
                1,
                |_| {
                    called.set(true);
                    1
                },
                |_| {
                    called.set(true);
                    None
                }
            ),
            Err(CounterError::InvalidLevel)
        );
        assert!(!called.get());
        assert!(state.counters.is_empty());
        assert!(state.started.is_empty());
    }

    #[test]
    fn overflow_does_not_commit_started_or_descendant_resets() {
        let mut state = CounterEngine::<u32, u32, u32>::default();
        let id = CounterIdentity::Shared(1);
        state
            .counters
            .insert(id, HashMap::from([(0, u32::MAX), (1, 7)]));
        let before = state.counters.clone();
        assert_eq!(
            state.advance(
                id,
                2,
                0,
                None,
                |_| 1,
                |level| Some(decimal(1, (level == 1).then_some(1)))
            ),
            Err(CounterError::Overflow)
        );
        assert_eq!(state.counters, before);
        assert!(state.started.is_empty());
    }

    fn marker(text: &str, counter: u32, format: &str) -> Result<String, CounterError> {
        CounterEngine::<u32, u32, u32>::default().resolve_text(
            CounterIdentity::Shared(1),
            0,
            counter,
            |_| 1,
            |_| {
                Some(LevelFacts {
                    start: 1,
                    restart: None,
                    format,
                    text,
                    legal: false,
                })
            },
        )
    }

    #[test]
    fn marker_budget_accepts_exact_size_and_rejects_one_over() {
        let exact = "x".repeat(MAX_MARKER_BYTES);
        assert_eq!(
            marker(&exact, 1, "decimal").unwrap().len(),
            MAX_MARKER_BYTES
        );
        let over = "x".repeat(MAX_MARKER_BYTES + 1);
        assert_eq!(
            marker(&over, 1, "decimal"),
            Err(CounterError::OutputTooLarge)
        );
    }

    #[test]
    fn repeated_placeholders_are_checked_before_replacement() {
        let fitting = "%1".repeat(MAX_MARKER_BYTES / 4);
        assert_eq!(
            marker(&fitting, 8, "upperRoman").unwrap().len(),
            MAX_MARKER_BYTES
        );
        let expanding = "%1".repeat(MAX_MARKER_BYTES / 4 + 1);
        assert_eq!(
            marker(&expanding, 8, "upperRoman"),
            Err(CounterError::OutputTooLarge)
        );
    }

    #[test]
    fn expanding_extreme_formats_reject_but_decimal_remains_small() {
        assert_eq!(
            marker("%1", u32::MAX, "upperLetter"),
            Err(CounterError::OutputTooLarge)
        );
        assert_eq!(
            marker("%1", u32::MAX, "upperRoman"),
            Err(CounterError::OutputTooLarge)
        );
        assert_eq!(marker("%1", u32::MAX, "decimal"), Ok(u32::MAX.to_string()));
    }
}
