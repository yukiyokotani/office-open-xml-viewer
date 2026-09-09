//! Compact paragraph-to-table context acquired from the shared DOC table grammar.
//! [MS-DOC] 2.4.3 and 2.4.6.6. This module does not apply table styles. Pre-existing
//! [`Properties`] allocations must be admitted by the caller; they are consumed and dropped here,
//! while the index accounts for its own additional fixed-size storage.

use super::{
    table::Properties,
    table_structure::{Assembler, Payload, RawEvent},
    unsupported, MAX_STORY_CONTROLS,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct ParagraphContext {
    pub(super) table_id: usize,
    pub(super) row_index: usize,
    pub(super) source_cell_index: Option<usize>,
    pub(super) ttp_id: usize,
}

#[derive(Debug, PartialEq, Eq)]
pub(super) struct TableContext {
    pub(super) rows: Vec<RowContext>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct RowContext {
    pub(super) ttp_id: usize,
    pub(super) source_cell_count: usize,
    pub(super) header: bool,
}

#[allow(dead_code)] // Wired into story projection by the subsequent style-resolution slice.
pub(super) struct Index {
    paragraphs: Vec<Option<ParagraphContext>>,
    tables: Vec<TableContext>,
}

#[allow(dead_code)] // Wired into story projection by the subsequent style-resolution slice.
impl Index {
    pub(super) fn build<I, A>(
        paragraph_count: usize,
        paragraphs: I,
        admit: &mut A,
    ) -> Result<Self, String>
    where
        I: IntoIterator<Item = (Properties, char)>,
        A: FnMut(usize) -> Result<(), String>,
    {
        if paragraph_count > MAX_STORY_CONTROLS + 1 {
            return Err(unsupported("Word table context paragraph budget exceeded"));
        }
        let mut contexts = allocate_none(paragraph_count, admit)?;
        let mut tables = Vec::new();
        let mut assembler = Assembler::<Ids>::new();
        let mut actual = 0usize;

        for (properties, mark) in paragraphs {
            if actual >= paragraph_count {
                return Err(unsupported("Word table context paragraph count mismatch"));
            }
            let depth = properties.depth()?;
            let payload = if depth == 0 {
                Ids::default()
            } else {
                Ids::one(actual, admit)?
            };
            assembler.push_raw(
                properties,
                mark,
                payload,
                Some(actual),
                |event, admit| index_event(event, &mut contexts, &mut tables, admit),
                admit,
            )?;
            actual += 1;
        }
        if actual != paragraph_count {
            return Err(unsupported("Word table context paragraph count mismatch"));
        }
        let _ = assembler.finish_raw(
            |event, admit| index_event(event, &mut contexts, &mut tables, admit),
            admit,
        )?;
        Ok(Self {
            paragraphs: contexts,
            tables,
        })
    }

    pub(super) fn paragraph(&self, id: usize) -> Result<Option<ParagraphContext>, String> {
        self.paragraphs
            .get(id)
            .copied()
            .ok_or_else(|| unsupported("Word table context paragraph outside index"))
    }

    pub(super) fn tables(&self) -> &[TableContext] {
        &self.tables
    }
}

#[derive(Default)]
struct Ids {
    own: Vec<usize>,
    nested: bool,
}

impl Ids {
    fn one<A: FnMut(usize) -> Result<(), String>>(
        id: usize,
        admit: &mut A,
    ) -> Result<Self, String> {
        let mut own = Vec::new();
        reserve_additional(&mut own, 1, admit)?;
        own.push(id);
        Ok(Self { own, nested: false })
    }
}

impl Payload for Ids {
    fn append<A: FnMut(usize) -> Result<(), String>>(
        &mut self,
        mut other: Self,
        admit: &mut A,
    ) -> Result<(), String> {
        reserve_additional(&mut self.own, other.own.len(), admit)?;
        self.own.append(&mut other.own);
        self.nested |= other.nested;
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.own.is_empty() && !self.nested
    }
}

fn index_event<A: FnMut(usize) -> Result<(), String>>(
    RawEvent(raw_tables): RawEvent<Ids>,
    contexts: &mut [Option<ParagraphContext>],
    tables: &mut Vec<TableContext>,
    admit: &mut A,
) -> Result<Ids, String> {
    let has_tables = !raw_tables.is_empty();
    for raw_table in raw_tables {
        let table_id = tables.len();
        let mut rows = Vec::new();
        reserve_additional(&mut rows, raw_table.rows.len(), admit)?;
        for (row_index, raw_row) in raw_table.rows.into_iter().enumerate() {
            let ttp_id = raw_row
                .ttp_id
                .ok_or_else(|| unsupported("Word table row lacks source paragraph"))?;
            set_innermost(
                contexts,
                ttp_id,
                ParagraphContext {
                    table_id,
                    row_index,
                    source_cell_index: None,
                    ttp_id,
                },
            )?;
            let source_cell_count = raw_row.source.cells.len();
            if raw_row.cells.len() != source_cell_count {
                return Err(unsupported("Word table context cell count mismatch"));
            }
            for (source_cell_index, payload) in raw_row.cells.into_iter().enumerate() {
                for id in payload.own {
                    set_innermost(
                        contexts,
                        id,
                        ParagraphContext {
                            table_id,
                            row_index,
                            source_cell_index: Some(source_cell_index),
                            ttp_id,
                        },
                    )?;
                }
            }
            rows.push(RowContext {
                ttp_id,
                source_cell_count,
                header: raw_row.source.header,
            });
        }
        reserve_additional(tables, 1, admit)?;
        tables.push(TableContext { rows });
    }
    Ok(Ids {
        own: Vec::new(),
        nested: has_tables,
    })
}

fn set_innermost(
    contexts: &mut [Option<ParagraphContext>],
    id: usize,
    context: ParagraphContext,
) -> Result<(), String> {
    let slot = contexts
        .get_mut(id)
        .ok_or_else(|| unsupported("Word table context paragraph outside index"))?;
    if slot.is_some() {
        return Err(unsupported("Word table context paragraph assigned twice"));
    }
    *slot = Some(context);
    Ok(())
}

fn allocate_none<A: FnMut(usize) -> Result<(), String>>(
    count: usize,
    admit: &mut A,
) -> Result<Vec<Option<ParagraphContext>>, String> {
    let mut values = Vec::new();
    reserve_additional(&mut values, count, admit)?;
    values.resize(count, None);
    Ok(values)
}

fn reserve_additional<T, A: FnMut(usize) -> Result<(), String>>(
    values: &mut Vec<T>,
    additional: usize,
    admit: &mut A,
) -> Result<(), String> {
    let required = values
        .len()
        .checked_add(additional)
        .ok_or("OUTPUT_TOO_LARGE")?;
    let old = values.capacity();
    if required <= old {
        return Ok(());
    }
    let doubled = old.checked_mul(2).unwrap_or(usize::MAX);
    let target = required.max(doubled).max(1);
    admit(
        target
            .checked_sub(old)
            .and_then(|growth| growth.checked_mul(std::mem::size_of::<T>()))
            .ok_or("OUTPUT_TOO_LARGE")?,
    )?;
    values
        .try_reserve_exact(target - values.len())
        .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
    admit(
        values
            .capacity()
            .saturating_sub(target)
            .checked_mul(std::mem::size_of::<T>())
            .ok_or("OUTPUT_TOO_LARGE")?,
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::{
        table::{Cell, Row},
        table_structure::{RawRow, RawTable},
    };

    fn cell(width: i32, flags: u16) -> Cell {
        Cell {
            width,
            flags,
            ..Cell::default()
        }
    }

    fn properties(depth: u32, row_end: bool, cells: Vec<Cell>) -> Properties {
        let mut value = Properties::default();
        value.apply(0x6649, &depth.to_le_bytes()).unwrap();
        value.row_end = row_end;
        value.row.cells = cells;
        value
    }

    fn build(paragraphs: Vec<(Properties, char)>) -> Index {
        Index::build(paragraphs.len(), paragraphs, &mut |_| Ok(())).unwrap()
    }

    #[test]
    fn indexes_multiple_rows_cells_and_ttp_contexts() {
        let mut first_end = properties(1, true, vec![cell(10, 0), cell(20, 0)]);
        first_end.row.header = true;
        let paragraphs = vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (first_end, '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(10, 0), cell(20, 0)]), '\u{7}'),
        ];
        let index = build(paragraphs);
        assert_eq!(index.tables().len(), 1);
        assert_eq!(index.tables()[0].rows.len(), 2);
        assert_eq!(
            index.tables()[0].rows,
            vec![
                RowContext {
                    ttp_id: 2,
                    source_cell_count: 2,
                    header: true,
                },
                RowContext {
                    ttp_id: 5,
                    source_cell_count: 2,
                    header: false,
                },
            ]
        );
        for (id, row, cell) in [
            (0, 0, Some(0)),
            (1, 0, Some(1)),
            (2, 0, None),
            (3, 1, Some(0)),
            (4, 1, Some(1)),
            (5, 1, None),
        ] {
            assert_eq!(
                index.paragraph(id).unwrap(),
                Some(ParagraphContext {
                    table_id: 0,
                    row_index: row,
                    source_cell_index: cell,
                    ttp_id: if row == 0 { 2 } else { 5 },
                })
            );
        }
    }

    #[test]
    fn indexes_thousands_of_paragraphs_in_one_source_cell() {
        let mut paragraphs = Vec::new();
        for _ in 0..4_096 {
            paragraphs.push((properties(1, false, vec![]), '\r'));
        }
        paragraphs.push((properties(1, false, vec![]), '\u{7}'));
        paragraphs.push((properties(1, true, vec![cell(10, 0)]), '\u{7}'));
        let index = build(paragraphs);

        for id in [0, 4_095, 4_096] {
            let context = index.paragraph(id).unwrap().unwrap();
            assert_eq!(context.source_cell_index, Some(0));
            assert_eq!(context.ttp_id, 4_097);
        }
        assert_eq!(
            index.paragraph(4_097).unwrap().unwrap().source_cell_index,
            None
        );
    }

    #[test]
    fn nested_table_context_is_innermost_and_outer_cell_keeps_placeholder() {
        let mut inner_cell = properties(2, false, vec![]);
        inner_cell.inner_cell = true;
        let mut inner_row = properties(2, false, vec![cell(10, 0)]);
        inner_row.inner_row = true;
        let paragraphs = vec![
            (inner_cell, '\r'),
            (inner_row, '\r'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(10, 0)]), '\u{7}'),
        ];
        let index = build(paragraphs);
        assert_eq!(index.tables().len(), 2);
        assert_eq!(index.paragraph(0).unwrap().unwrap().table_id, 0);
        assert_eq!(index.paragraph(1).unwrap().unwrap().table_id, 0);
        assert_eq!(index.paragraph(2).unwrap().unwrap().table_id, 1);
        assert_eq!(index.paragraph(3).unwrap().unwrap().table_id, 1);
        assert_eq!(index.paragraph(0).unwrap().unwrap().ttp_id, 1);
        assert_eq!(index.paragraph(2).unwrap().unwrap().ttp_id, 3);
    }

    #[test]
    fn source_merge_continuations_keep_distinct_paragraph_contexts() {
        let index = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(10, 2), cell(10, 1)]), '\u{7}'),
        ]);
        assert_eq!(
            index.paragraph(0).unwrap().unwrap().source_cell_index,
            Some(0)
        );
        assert_eq!(
            index.paragraph(1).unwrap().unwrap().source_cell_index,
            Some(1)
        );
        assert_eq!(index.tables()[0].rows[0].source_cell_count, 2);
    }

    #[test]
    fn table_identity_changes_create_stable_table_ids() {
        let mut first = properties(1, true, vec![cell(10, 0)]);
        first
            .row
            .identity
            .insert(0x563a, 7u16.to_le_bytes().to_vec());
        let mut second = properties(1, true, vec![cell(10, 0)]);
        second
            .row
            .identity
            .insert(0x563a, 8u16.to_le_bytes().to_vec());
        let index = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (first, '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (second, '\u{7}'),
        ]);
        assert_eq!(index.tables().len(), 2);
        assert_eq!(index.paragraph(0).unwrap().unwrap().table_id, 0);
        assert_eq!(index.paragraph(2).unwrap().unwrap().table_id, 1);
    }

    #[test]
    fn empty_and_non_table_stories_have_no_contexts() {
        let empty = Index::build(0, Vec::new(), &mut |_| Ok(())).unwrap();
        assert!(empty.tables().is_empty());
        assert!(empty.paragraph(0).is_err());
        let plain = build(vec![(Properties::default(), '\r')]);
        assert!(plain.tables().is_empty());
        assert_eq!(plain.paragraph(0).unwrap(), None);
        assert!(plain.paragraph(1).is_err());
    }

    #[test]
    fn rejects_paragraph_count_and_table_grammar_mismatches() {
        for declared in [0, 2] {
            let error = Index::build(declared, vec![(Properties::default(), '\r')], &mut |_| {
                Ok(())
            })
            .err()
            .unwrap();
            assert!(error.contains("paragraph count mismatch"), "{error}");
        }
        let error = Index::build(
            1,
            vec![(properties(1, true, vec![cell(10, 0)]), '\u{7}')],
            &mut |_| Ok(()),
        )
        .err()
        .unwrap();
        assert!(
            error.contains("row definition does not match cell marks"),
            "{error}"
        );

        let mut inner_cell = properties(2, false, vec![]);
        inner_cell.inner_cell = true;
        let mut inner_row = properties(2, false, vec![cell(10, 0)]);
        inner_row.inner_row = true;
        let error = Index::build(2, vec![(inner_cell, '\r'), (inner_row, '\r')], &mut |_| {
            Ok(())
        })
        .err()
        .unwrap();
        assert!(error.contains("unterminated Word table row"), "{error}");
    }

    #[test]
    fn budget_denial_precedes_index_allocation_and_iterator_consumption() {
        let consumed = std::cell::Cell::new(0);
        let paragraphs = std::iter::from_fn(|| {
            consumed.set(consumed.get() + 1);
            Some((Properties::default(), '\r'))
        });
        assert_eq!(
            Index::build(1, paragraphs, &mut |_| Err("denied".into()))
                .err()
                .as_deref(),
            Some("denied")
        );
        assert_eq!(consumed.get(), 0);
        assert!(
            Index::build(MAX_STORY_CONTROLS + 2, Vec::new(), &mut |_| panic!(
                "budget must be rejected before allocation"
            ))
            .err()
            .unwrap()
            .contains("paragraph budget")
        );
    }

    fn nested(depth: u32) -> Vec<(Properties, char)> {
        let mut paragraphs = Vec::new();
        for level in 1..=depth {
            paragraphs.push((properties(level, false, vec![]), '\r'));
        }
        for level in (1..=depth).rev() {
            let mut cell_end = properties(level, false, vec![]);
            let mut row_end = properties(level, false, vec![cell(10, 0)]);
            let mark = if level == 1 { '\u{7}' } else { '\r' };
            if level == 1 {
                row_end.row_end = true;
            } else {
                cell_end.inner_cell = true;
                row_end.inner_row = true;
            }
            paragraphs.push((cell_end, mark));
            paragraphs.push((row_end, mark));
        }
        paragraphs
    }

    #[test]
    fn deep_nesting_maps_each_id_once_without_ancestor_fanout() {
        let paragraphs = nested(32);
        let thirty_two = Index::build(paragraphs.len(), paragraphs, &mut |_| Ok(())).unwrap();
        assert_eq!(thirty_two.tables().len(), 32);
        assert!((0..96).all(|id| thirty_two.paragraph(id).unwrap().is_some()));
    }

    #[test]
    fn reserve_growth_is_geometric_for_large_cells_and_table_lists() {
        let mut paragraph_ids = Vec::new();
        let mut id_admissions = 0usize;
        for id in 0..4_096 {
            reserve_additional(&mut paragraph_ids, 1, &mut |bytes| {
                id_admissions += usize::from(bytes != 0);
                Ok(())
            })
            .unwrap();
            paragraph_ids.push(id);
        }
        let mut tables = Vec::new();
        let mut table_admissions = 0usize;
        for _ in 0..4_096 {
            reserve_additional(&mut tables, 1, &mut |bytes| {
                table_admissions += usize::from(bytes != 0);
                Ok(())
            })
            .unwrap();
            tables.push(TableContext { rows: Vec::new() });
        }

        // Doubling from one element needs 1 + log2(count) growths. Each growth
        // has at most one requested-capacity and one allocator-excess admission.
        let maximum_admissions = 2 * (1 + 4_096usize.ilog2() as usize);
        assert!(id_admissions <= maximum_admissions);
        assert!(table_admissions <= maximum_admissions);
    }

    #[test]
    fn empty_raw_event_does_not_fabricate_nested_content() {
        let mut contexts = Vec::new();
        let mut tables = Vec::new();
        let payload = index_event(
            RawEvent(Vec::new()),
            &mut contexts,
            &mut tables,
            &mut |_| Ok(()),
        )
        .unwrap();
        assert!(payload.is_empty());
        assert!(tables.is_empty());
    }

    #[test]
    fn nonempty_raw_event_returns_only_a_nested_placeholder() {
        let mut source = Row::default();
        source.cells = vec![cell(10, 0)];
        let raw = RawEvent(vec![RawTable {
            rows: vec![RawRow {
                source,
                cells: vec![Ids {
                    own: vec![0],
                    nested: false,
                }],
                ttp_id: Some(1),
            }],
        }]);
        let mut contexts = vec![None; 2];
        let mut tables = Vec::new();
        let payload = index_event(raw, &mut contexts, &mut tables, &mut |_| Ok(())).unwrap();
        assert!(payload.own.is_empty());
        assert!(payload.nested);
        assert!(contexts.into_iter().all(|context| context.is_some()));
    }

    #[test]
    fn duplicate_context_assignment_is_rejected() {
        let context = ParagraphContext {
            table_id: 0,
            row_index: 0,
            source_cell_index: Some(0),
            ttp_id: 1,
        };
        let mut contexts = vec![Some(context)];
        let error = set_innermost(&mut contexts, 0, context).unwrap_err();
        assert!(error.contains("assigned twice"), "{error}");
    }
}
