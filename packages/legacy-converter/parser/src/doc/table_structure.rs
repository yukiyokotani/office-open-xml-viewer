//! Representation-independent MS-DOC table grammar and grid planning.
//!
//! [MS-DOC] 2.4.3 encodes cells and rows with paragraph marks. This module is
//! the single owner of that nesting grammar and of the union-grid construction
//! used by both the legacy DOCX serializer and the direct model producer.

use std::borrow::Borrow;

use super::{
    table::{Cell, Properties, Row},
    unsupported,
};

pub(super) trait Payload: Default {
    fn append<A: FnMut(usize) -> Result<(), String>>(
        &mut self,
        other: Self,
        admit: &mut A,
    ) -> Result<(), String>;
    fn is_empty(&self) -> bool;
}

struct Pending<P, R> {
    rows: Vec<RawRow<P, R>>,
    cells: Vec<P>,
    cell: P,
}

impl<P: Default, R> Default for Pending<P, R> {
    fn default() -> Self {
        Self {
            rows: Vec::new(),
            cells: Vec::new(),
            cell: P::default(),
        }
    }
}

pub(super) struct Assembler<P, R = Row> {
    stack: Vec<Pending<P, R>>,
    body: P,
    rows: usize,
}

pub(super) struct Event<P>(pub(super) Vec<LogicalTable<P>>);

/// Grammar-complete logical tables before union-grid planning and merge
/// projection. Ownership lets a raw consumer avoid geometry planning entirely;
/// planned consumers move the same rows onward without cloning them.
pub(super) struct RawEvent<P, R = Row>(pub(super) Vec<RawTable<P, R>>);

pub(super) struct RawTable<P, R = Row> {
    pub(super) rows: Vec<RawRow<P, R>>,
}

pub(super) struct RawRow<P, R = Row> {
    pub(super) source: R,
    pub(super) cells: Vec<P>,
    #[allow(dead_code)] // Consumed by the subsequent table-context indexing slice.
    pub(super) ttp_id: Option<usize>,
}

pub(super) struct LogicalTable<P> {
    pub(super) grid: Vec<i32>,
    pub(super) origin: i32,
    pub(super) total: i32,
    pub(super) rows: Vec<PlannedRow<P>>,
}

pub(super) struct PlannedRow<P> {
    pub(super) source: Row,
    pub(super) grid_before: usize,
    pub(super) width_before: i32,
    pub(super) grid_after: usize,
    pub(super) width_after: i32,
    pub(super) is_header: bool,
    pub(super) source_cell_count: usize,
    pub(super) cells: Vec<PlannedCell<P>>,
}

pub(super) struct PlannedCell<P> {
    pub(super) source: Cell,
    pub(super) source_index: usize,
    pub(super) source_end: usize,
    pub(super) width: i32,
    pub(super) grid_span: usize,
    pub(super) vertical: u16,
    pub(super) content: P,
}

impl<P: Payload, R: Borrow<Row>> Assembler<P, R> {
    pub(super) fn new() -> Self {
        Self {
            stack: Vec::new(),
            body: P::default(),
            rows: 0,
        }
    }

    pub(super) fn push_raw<E, A>(
        &mut self,
        props: Properties<R>,
        mark: char,
        paragraph: P,
        // Retained only when this paragraph is the row's TTP mark.
        source_id: Option<usize>,
        mut emit: E,
        admit: &mut A,
    ) -> Result<bool, String>
    where
        E: FnMut(RawEvent<P, R>, &mut A) -> Result<P, String>,
        A: FnMut(usize) -> Result<(), String>,
    {
        let depth = props.depth()?;
        while self.stack.len() > depth {
            self.close(&mut emit, admit)?;
        }
        while self.stack.len() < depth {
            reserve_one(&mut self.stack, admit)?;
            self.stack.push(Pending::default());
        }
        if depth == 0 {
            self.body.append(paragraph, admit)?;
            return Ok(true);
        }
        let row_end = if depth == 1 {
            mark == '\u{7}' && props.row_end
        } else {
            mark == '\r' && props.inner_row
        };
        let cell_end = if depth == 1 {
            mark == '\u{7}'
        } else {
            mark == '\r' && props.inner_cell
        };
        let current = self.stack.last_mut().expect("open table");
        if row_end {
            self.rows = self.rows.checked_add(1).ok_or("OUTPUT_TOO_LARGE")?;
            if self.rows > 100_000 {
                return Err(unsupported("Word table row budget exceeded"));
            }
            if !current.cell.is_empty()
                || current.cells.len() != props.row.borrow().cells.len()
                || current.cells.is_empty()
            {
                return Err(unsupported("Word row definition does not match cell marks"));
            }
            reserve_one(&mut current.rows, admit)?;
            current.rows.push(RawRow {
                source: props.row,
                cells: std::mem::take(&mut current.cells),
                ttp_id: source_id,
            });
        } else {
            current.cell.append(paragraph, admit)?;
            if cell_end {
                if current.cells.len() >= 63 {
                    return Err(unsupported("too many Word table cell marks"));
                }
                reserve_one(&mut current.cells, admit)?;
                current.cells.push(std::mem::take(&mut current.cell));
            }
        }
        Ok(!row_end)
    }

    fn close<E, A>(&mut self, emit: &mut E, admit: &mut A) -> Result<(), String>
    where
        E: FnMut(RawEvent<P, R>, &mut A) -> Result<P, String>,
        A: FnMut(usize) -> Result<(), String>,
    {
        let pending = self.stack.pop().expect("open table");
        if !pending.cell.is_empty() || !pending.cells.is_empty() {
            return Err(unsupported("unterminated Word table row"));
        }
        let mut tables = Vec::new();
        let mut rows = pending.rows.into_iter().peekable();
        while let Some(first) = rows.next() {
            let bidi = first.source.borrow().bidi;
            let mut group = Vec::new();
            reserve_one(&mut group, admit)?;
            group.push(first);
            while rows.peek().is_some_and(|row| {
                row.source.borrow().bidi == bidi
                    && row.source.borrow().identity == group[0].source.borrow().identity
            }) {
                reserve_one(&mut group, admit)?;
                group.push(rows.next().expect("peeked row"));
            }
            reserve_one(&mut tables, admit)?;
            tables.push(RawTable { rows: group });
        }
        let payload = emit(RawEvent(tables), admit)?;
        if let Some(parent) = self.stack.last_mut() {
            parent.cell.append(payload, admit)?;
        } else {
            self.body.append(payload, admit)?;
        }
        Ok(())
    }

    pub(super) fn finish_raw<E, A>(mut self, mut emit: E, admit: &mut A) -> Result<P, String>
    where
        E: FnMut(RawEvent<P, R>, &mut A) -> Result<P, String>,
        A: FnMut(usize) -> Result<(), String>,
    {
        while !self.stack.is_empty() {
            self.close(&mut emit, admit)?;
        }
        Ok(self.body)
    }

    #[cfg(test)]
    pub(super) fn set_row_count(&mut self, rows: usize) {
        self.rows = rows;
    }
}

impl<P: Payload> Assembler<P> {
    pub(super) fn push<F, A>(
        &mut self,
        props: Properties,
        mark: char,
        paragraph: P,
        mut emit: F,
        admit: &mut A,
    ) -> Result<bool, String>
    where
        F: FnMut(Event<P>) -> Result<P, String>,
        A: FnMut(usize) -> Result<(), String>,
    {
        self.push_raw(
            props,
            mark,
            paragraph,
            None,
            |raw, admit| emit(plan_event(raw, admit)?),
            admit,
        )
    }

    pub(super) fn finish<F, A>(self, mut emit: F, admit: &mut A) -> Result<P, String>
    where
        F: FnMut(Event<P>) -> Result<P, String>,
        A: FnMut(usize) -> Result<(), String>,
    {
        self.finish_raw(|raw, admit| emit(plan_event(raw, admit)?), admit)
    }
}

fn plan_event<P: Default, A: FnMut(usize) -> Result<(), String>>(
    RawEvent(raw_tables): RawEvent<P>,
    admit: &mut A,
) -> Result<Event<P>, String> {
    let mut tables = Vec::new();
    reserve_exact(&mut tables, raw_tables.len(), admit)?;
    for raw in raw_tables {
        tables.push(plan(raw, admit)?);
    }
    Ok(Event(tables))
}

fn plan<P: Default, A: FnMut(usize) -> Result<(), String>>(
    raw: RawTable<P>,
    admit: &mut A,
) -> Result<LogicalTable<P>, String> {
    let rows = raw.rows;
    let (boundaries, _) = grid_boundaries(&rows, admit)?;
    let grid_len = boundaries.iter().try_fold(0usize, |sum, (_, count)| {
        sum.checked_add(*count).ok_or("OUTPUT_TOO_LARGE")
    })?;
    if grid_len > 65_536 {
        return Err(unsupported("Word table grid budget exceeded"));
    }
    let mut grid = Vec::new();
    reserve_exact(&mut grid, grid_len, admit)?;
    for (edge, count) in boundaries {
        grid.extend(std::iter::repeat_n(edge, count));
    }
    if grid.len() < 2 {
        return Err(unsupported("Word table has no cell boundaries"));
    }
    let origin = grid[0];
    let total = grid[grid.len() - 1]
        .checked_sub(origin)
        .ok_or("OUTPUT_TOO_LARGE")?;
    let mut header_prefix = true;
    let mut planned_rows = Vec::new();
    reserve_exact(&mut planned_rows, rows.len(), admit)?;
    for RawRow {
        source: mut row,
        cells: contents,
        ttp_id: _,
    } in rows
    {
        let edge = row.origin();
        let before = grid.partition_point(|value| *value < edge);
        let mut cell_grid = Vec::new();
        reserve_exact(
            &mut cell_grid,
            row.cells.len().checked_add(1).ok_or("OUTPUT_TOO_LARGE")?,
            admit,
        )?;
        cell_grid.push(before);
        let mut endpoint = edge;
        for cell in &row.cells {
            endpoint = endpoint.checked_add(cell.width).ok_or("OUTPUT_TOO_LARGE")?;
            let next = if cell.width == 0 {
                cell_grid
                    .last()
                    .copied()
                    .unwrap()
                    .checked_add(1)
                    .ok_or("OUTPUT_TOO_LARGE")?
            } else {
                grid.partition_point(|value| *value < endpoint)
            };
            cell_grid.push(next);
        }
        let after = grid.len() - 1 - cell_grid.last().copied().unwrap();
        header_prefix &= row.header;
        let source_cells = std::mem::take(&mut row.cells);
        let source_cell_count = source_cells.len();
        let mut cells = Vec::new();
        reserve_exact(&mut cells, source_cells.len(), admit)?;
        let mut entries = source_cells
            .into_iter()
            .zip(contents)
            .enumerate()
            .peekable();
        while let Some((i, (source, content))) = entries.next() {
            let mut end = i + 1;
            let mut width = source.width;
            if source.flags & 3 >= 2 {
                while entries
                    .peek()
                    .is_some_and(|(_, (cell, _))| cell.flags & 3 == 1)
                {
                    let (_, (continued, _discarded_content)) = entries.next().expect("peeked cell");
                    width = width
                        .checked_add(continued.width)
                        .ok_or("OUTPUT_TOO_LARGE")?;
                    end += 1;
                }
            }
            let vertical = (source.flags >> 5) & 3;
            cells.push(PlannedCell {
                source,
                source_index: i,
                source_end: end,
                width,
                grid_span: cell_grid[end] - cell_grid[i],
                vertical,
                content,
            });
        }
        planned_rows.push(PlannedRow {
            grid_before: before,
            width_before: edge.checked_sub(origin).ok_or("OUTPUT_TOO_LARGE")?,
            grid_after: after,
            width_after: grid[grid.len() - 1]
                .checked_sub(endpoint)
                .ok_or("OUTPUT_TOO_LARGE")?,
            is_header: header_prefix,
            source_cell_count,
            source: row,
            cells,
        });
    }
    Ok(LogicalTable {
        grid,
        origin,
        total,
        rows: planned_rows,
    })
}

/// Collect every row edge into one pre-admitted scratch vector, sort once, then
/// coalesce `(edge,row)` groups to the maximum duplicate multiplicity observed
/// in any row. Work is O(edges log edges); retained union size remains capped.
fn grid_boundaries<P, A: FnMut(usize) -> Result<(), String>>(
    rows: &[RawRow<P>],
    admit: &mut A,
) -> Result<(Vec<(i32, usize)>, usize), String> {
    let edge_count = rows.iter().try_fold(0usize, |sum, raw| {
        sum.checked_add(
            raw.source
                .cells
                .len()
                .checked_add(1)
                .ok_or("OUTPUT_TOO_LARGE")?,
        )
        .ok_or("OUTPUT_TOO_LARGE")
    })?;
    let mut boundaries = Vec::<(i32, usize)>::new();
    reserve_exact(&mut boundaries, edge_count, admit)?;
    for (row_index, raw) in rows.iter().enumerate() {
        let row = &raw.source;
        let mut edge = row.origin();
        boundaries.push((edge, row_index));
        for cell in &row.cells {
            edge = edge.checked_add(cell.width).ok_or("OUTPUT_TOO_LARGE")?;
            boundaries.push((edge, row_index));
        }
    }
    #[cfg(test)]
    let comparisons = std::cell::Cell::new(0usize);
    boundaries.sort_unstable_by(|left, right| {
        #[cfg(test)]
        comparisons.set(comparisons.get().saturating_add(1));
        left.cmp(right)
    });
    let mut read = 0;
    let mut write = 0;
    while read < boundaries.len() {
        let edge = boundaries[read].0;
        let mut maximum = 0;
        while read < boundaries.len() && boundaries[read].0 == edge {
            let row = boundaries[read].1;
            let start = read;
            while read < boundaries.len() && boundaries[read] == (edge, row) {
                read += 1;
            }
            maximum = maximum.max(read - start);
        }
        boundaries[write] = (edge, maximum);
        write += 1;
        if write > 65_536 {
            return Err(unsupported("Word table grid budget exceeded"));
        }
    }
    boundaries.truncate(write);
    #[cfg(test)]
    let comparison_count = comparisons.get();
    #[cfg(not(test))]
    let comparison_count = 0;
    Ok((boundaries, comparison_count))
}

fn reserve_one<T, A: FnMut(usize) -> Result<(), String>>(
    values: &mut Vec<T>,
    admit: &mut A,
) -> Result<(), String> {
    if values.len() == values.capacity() {
        let old = values.capacity();
        let growth = old.max(1);
        admit(
            growth
                .checked_mul(std::mem::size_of::<T>())
                .ok_or("OUTPUT_TOO_LARGE")?,
        )?;
        values
            .try_reserve_exact(growth)
            .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
        let expected = old.checked_add(growth).ok_or("OUTPUT_TOO_LARGE")?;
        admit(
            values
                .capacity()
                .saturating_sub(expected)
                .checked_mul(std::mem::size_of::<T>())
                .ok_or("OUTPUT_TOO_LARGE")?,
        )?;
    }
    Ok(())
}

fn reserve_exact<T, A: FnMut(usize) -> Result<(), String>>(
    values: &mut Vec<T>,
    count: usize,
    admit: &mut A,
) -> Result<(), String> {
    admit(
        count
            .checked_mul(std::mem::size_of::<T>())
            .ok_or("OUTPUT_TOO_LARGE")?,
    )?;
    values
        .try_reserve_exact(count)
        .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
    admit(
        values
            .capacity()
            .saturating_sub(count)
            .checked_mul(std::mem::size_of::<T>())
            .ok_or("OUTPUT_TOO_LARGE")?,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    impl Payload for Vec<usize> {
        fn append<A: FnMut(usize) -> Result<(), String>>(
            &mut self,
            mut other: Self,
            admit: &mut A,
        ) -> Result<(), String> {
            reserve_exact(self, other.len(), admit)?;
            self.append(&mut other);
            Ok(())
        }

        fn is_empty(&self) -> bool {
            Vec::is_empty(self)
        }
    }

    fn cell(width: i32) -> Cell {
        Cell {
            width,
            ..Cell::default()
        }
    }

    fn raw<P>(source: Row, cells: Vec<P>) -> RawRow<P> {
        RawRow {
            source,
            cells,
            ttp_id: None,
        }
    }

    fn properties(depth: u32, row_end: bool, cells: Vec<Cell>) -> Properties {
        let mut value = Properties::default();
        value.apply(0x6649, &depth.to_le_bytes()).unwrap();
        value.row_end = row_end;
        value.row.cells = cells;
        value
    }

    #[test]
    fn raw_event_retains_continuation_payload_and_ttp_identifier() {
        let mut assembler = Assembler::<Vec<usize>>::new();
        fn admit(_: usize) -> Result<(), String> {
            Ok(())
        }
        for id in [10, 11] {
            assembler
                .push_raw(
                    properties(1, false, vec![]),
                    '\u{7}',
                    vec![id],
                    None,
                    |_, _| unreachable!(),
                    &mut admit,
                )
                .unwrap();
        }
        let mut first = cell(10);
        first.flags = 2;
        let mut continuation = cell(10);
        continuation.flags = 1;
        assembler
            .push_raw(
                properties(1, true, vec![first, continuation]),
                '\u{7}',
                vec![],
                Some(99),
                |_, _| unreachable!(),
                &mut admit,
            )
            .unwrap();
        let mut observed = None;
        let body = assembler
            .finish_raw(
                |RawEvent(mut tables), _| {
                    assert_eq!(tables.len(), 1);
                    let row = &tables[0].rows[0];
                    observed = Some((row.cells.clone(), row.ttp_id));
                    Ok(tables
                        .pop()
                        .unwrap()
                        .rows
                        .into_iter()
                        .flat_map(|row| row.cells.into_iter().flatten())
                        .collect())
                },
                &mut admit,
            )
            .unwrap();
        assert_eq!(observed, Some((vec![vec![10], vec![11]], Some(99))));
        assert_eq!(body, vec![10, 11]);
    }

    #[test]
    fn raw_nested_events_return_inner_payload_to_outer_owner() {
        let mut assembler = Assembler::<Vec<usize>>::new();
        fn admit(_: usize) -> Result<(), String> {
            Ok(())
        }
        assembler
            .push_raw(
                properties(1, false, vec![]),
                '\r',
                vec![1],
                None,
                |_, _| unreachable!(),
                &mut admit,
            )
            .unwrap();
        let mut inner_cell_end = properties(2, false, vec![]);
        inner_cell_end.inner_cell = true;
        assembler
            .push_raw(
                inner_cell_end,
                '\r',
                vec![2],
                None,
                |_, _| unreachable!(),
                &mut admit,
            )
            .unwrap();
        assembler
            .push_raw(
                properties(2, false, vec![]),
                '\r',
                vec![],
                None,
                |_, _| unreachable!(),
                &mut admit,
            )
            .unwrap();
        let mut inner_end = properties(2, false, vec![cell(10)]);
        inner_end.inner_row = true;
        assembler
            .push_raw(
                inner_end,
                '\r',
                vec![],
                Some(20),
                |_, _| unreachable!(),
                &mut admit,
            )
            .unwrap();
        let mut seen = Vec::new();
        assembler
            .push_raw(
                properties(1, false, vec![]),
                '\u{7}',
                vec![],
                None,
                |RawEvent(tables), _| {
                    let ids: Vec<_> = tables[0].rows[0].cells.iter().flatten().copied().collect();
                    seen.push((ids.clone(), tables[0].rows[0].ttp_id));
                    Ok(ids)
                },
                &mut admit,
            )
            .unwrap();
        assembler
            .push_raw(
                properties(1, true, vec![cell(10)]),
                '\u{7}',
                vec![],
                Some(10),
                |_, _| unreachable!(),
                &mut admit,
            )
            .unwrap();
        let body = assembler
            .finish_raw(
                |RawEvent(tables), _| {
                    let ids: Vec<_> = tables[0].rows[0].cells.iter().flatten().copied().collect();
                    seen.push((ids.clone(), tables[0].rows[0].ttp_id));
                    Ok(ids)
                },
                &mut admit,
            )
            .unwrap();
        assert_eq!(seen, vec![(vec![2], Some(20)), (vec![1, 2], Some(10))]);
        assert_eq!(body, vec![1, 2]);
    }

    #[test]
    fn raw_event_preserves_logical_table_splits() {
        let mut assembler = Assembler::<Vec<usize>>::new();
        fn admit(_: usize) -> Result<(), String> {
            Ok(())
        }
        for (id, identity) in [(1, 7), (2, 8)] {
            assembler
                .push_raw(
                    properties(1, false, vec![]),
                    '\u{7}',
                    vec![id],
                    None,
                    |_, _| unreachable!(),
                    &mut admit,
                )
                .unwrap();
            let mut end = properties(1, true, vec![cell(10)]);
            end.row
                .identity
                .insert(0x563a, (identity as u16).to_le_bytes().to_vec());
            assembler
                .push_raw(
                    end,
                    '\u{7}',
                    vec![],
                    Some(id),
                    |_, _| unreachable!(),
                    &mut admit,
                )
                .unwrap();
        }
        assembler
            .finish_raw(
                |RawEvent(tables), _| {
                    assert_eq!(tables.len(), 2);
                    assert_eq!(tables[0].rows[0].ttp_id, Some(1));
                    assert_eq!(tables[1].rows[0].ttp_id, Some(2));
                    Ok(vec![])
                },
                &mut admit,
            )
            .unwrap();
    }

    #[test]
    fn raw_path_preserves_row_grammar_errors() {
        let mut assembler = Assembler::<Vec<usize>>::new();
        let error = assembler
            .push_raw(
                properties(1, true, vec![cell(10)]),
                '\u{7}',
                vec![],
                Some(1),
                |_, _| unreachable!(),
                &mut |_| Ok(()),
            )
            .unwrap_err();
        assert_eq!(
            error,
            "UNSUPPORTED:Word row definition does not match cell marks"
        );
    }

    #[test]
    fn raw_event_metadata_is_admitted_before_emission() {
        fn seed() -> Assembler<Vec<usize>> {
            let mut assembler = Assembler::new();
            assembler
                .push_raw(
                    properties(1, false, vec![]),
                    '\u{7}',
                    vec![1],
                    None,
                    |_, _| unreachable!(),
                    &mut |_| Ok(()),
                )
                .unwrap();
            assembler
                .push_raw(
                    properties(1, true, vec![cell(10)]),
                    '\u{7}',
                    vec![],
                    Some(2),
                    |_, _| unreachable!(),
                    &mut |_| Ok(()),
                )
                .unwrap();
            assembler
        }

        let mut admissions = Vec::new();
        let emitted = std::cell::Cell::new(false);
        seed()
            .finish_raw(
                |_, _| {
                    emitted.set(true);
                    Ok(vec![])
                },
                &mut |bytes| {
                    if !emitted.get() {
                        admissions.push(bytes);
                    }
                    Ok(())
                },
            )
            .unwrap();
        assert!(!admissions.is_empty());
        for denied in 0..admissions.len() {
            let mut call = 0;
            let mut emitted = false;
            let error = seed()
                .finish_raw(
                    |_, _| {
                        emitted = true;
                        Ok(vec![])
                    },
                    &mut |_| {
                        let current = call;
                        call += 1;
                        if current == denied {
                            Err("denied".into())
                        } else {
                            Ok(())
                        }
                    },
                )
                .unwrap_err();
            assert_eq!(error, "denied");
            assert!(!emitted, "callback ran before admission {denied}");
        }
    }

    #[test]
    fn raw_callback_error_propagates_without_returning_partial_body() {
        let mut assembler = Assembler::<Vec<usize>>::new();
        assembler
            .push_raw(
                properties(1, false, vec![]),
                '\u{7}',
                vec![1],
                None,
                |_, _| unreachable!(),
                &mut |_| Ok(()),
            )
            .unwrap();
        assembler
            .push_raw(
                properties(1, true, vec![cell(10)]),
                '\u{7}',
                vec![],
                Some(2),
                |_, _| unreachable!(),
                &mut |_| Ok(()),
            )
            .unwrap();
        assert_eq!(
            assembler
                .finish_raw(|_, _| Err("callback".into()), &mut |_| Ok(()))
                .unwrap_err(),
            "callback"
        );
    }

    #[test]
    fn raw_path_skips_union_grid_planning_but_planned_wrapper_does_not() {
        fn feed(mut assembler: Assembler<Vec<usize>>, raw: bool) -> Result<Vec<usize>, String> {
            fn admit(_: usize) -> Result<(), String> {
                Ok(())
            }
            for row_index in 0..1025 {
                for _ in 0..63 {
                    if raw {
                        assembler.push_raw(
                            properties(1, false, vec![]),
                            '\u{7}',
                            vec![],
                            None,
                            |_, _| unreachable!(),
                            &mut admit,
                        )?;
                    } else {
                        assembler.push(
                            properties(1, false, vec![]),
                            '\u{7}',
                            vec![],
                            |Event(_)| Ok(vec![]),
                            &mut admit,
                        )?;
                    }
                }
                let mut end = properties(1, true, vec![cell(0); 63]);
                end.row.left = row_index;
                if raw {
                    assembler.push_raw(
                        end,
                        '\u{7}',
                        vec![],
                        Some(row_index as usize),
                        |_, _| unreachable!(),
                        &mut admit,
                    )?;
                } else {
                    assembler.push(end, '\u{7}', vec![], |Event(_)| Ok(vec![]), &mut admit)?;
                }
            }
            if raw {
                assembler.finish_raw(
                    |RawEvent(tables), _| {
                        assert_eq!(tables.len(), 1);
                        assert_eq!(tables[0].rows.len(), 1025);
                        assert!(tables[0].rows.iter().all(|row| row.cells.len() == 63));
                        Ok(vec![])
                    },
                    &mut admit,
                )
            } else {
                assembler.finish(|Event(_)| Ok(vec![]), &mut admit)
            }
        }

        assert!(feed(Assembler::new(), true).is_ok());
        assert_eq!(
            feed(Assembler::new(), false).unwrap_err(),
            "UNSUPPORTED:Word table grid budget exceeded"
        );
    }

    #[test]
    fn reverse_ordered_distinct_edges_sort_once_with_bounded_comparisons() {
        let row_count = 32_768usize;
        let mut rows = Vec::<RawRow<()>>::with_capacity(row_count);
        for index in 0..row_count {
            let mut row = Row::default();
            row.left = i32::try_from((row_count - index) * 2).unwrap();
            row.cells.push(cell(1));
            rows.push(raw(row, Vec::new()));
        }
        let mut admitted = 0usize;
        let (boundaries, comparisons) = grid_boundaries(&rows, &mut |bytes| {
            admitted = admitted.checked_add(bytes).ok_or("OUTPUT_TOO_LARGE")?;
            Ok(())
        })
        .unwrap();
        assert_eq!(boundaries.len(), 65_536);
        assert_eq!(boundaries.first(), Some(&(2, 1)));
        assert_eq!(boundaries.last(), Some(&(65_537, 1)));
        assert!(boundaries.windows(2).all(|pair| pair[0].0 < pair[1].0));
        let edges = row_count * 2;
        // This deterministic bound is twice log2(N) comparisons per entry;
        // repeated sorted insertion exceeds it by orders of magnitude.
        assert!(comparisons < edges * 32);
        assert!(admitted >= edges * std::mem::size_of::<(i32, usize)>());
    }

    #[test]
    fn duplicate_edge_multiplicity_is_the_max_per_row_not_the_total() {
        let mut first = Row::default();
        first.cells = vec![cell(0), cell(0), cell(10)];
        let mut second = Row::default();
        second.cells = vec![cell(0), cell(10)];
        let mut admitted = 0usize;
        let (boundaries, _) = grid_boundaries(
            &[raw(first, Vec::<()>::new()), raw(second, Vec::new())],
            &mut |bytes| {
                admitted += bytes;
                Ok(())
            },
        )
        .unwrap();
        assert_eq!(boundaries, vec![(0, 3), (10, 1)]);
        assert!(admitted > 0);
    }

    #[test]
    fn expanded_duplicate_grid_still_obeys_the_65536_budget() {
        let mut row = Row::default();
        row.cells = (0..65_536).map(|_| cell(0)).collect();
        assert_eq!(
            plan(
                RawTable {
                    rows: vec![raw(row, vec![(); 65_536])],
                },
                &mut |_| Ok(()),
            )
            .err()
            .as_deref(),
            Some("UNSUPPORTED:Word table grid budget exceeded")
        );
    }

    #[test]
    fn scratch_is_admitted_before_allocation_or_population() {
        let mut row = Row::default();
        row.cells = vec![cell(10), cell(20)];
        let expected = 3 * std::mem::size_of::<(i32, usize)>();
        let mut calls = Vec::new();
        let error = grid_boundaries(&[raw(row, Vec::<()>::new())], &mut |bytes| {
            calls.push(bytes);
            Err("denied".into())
        })
        .unwrap_err();
        assert_eq!(error, "denied");
        assert_eq!(calls, vec![expected]);
    }
}
