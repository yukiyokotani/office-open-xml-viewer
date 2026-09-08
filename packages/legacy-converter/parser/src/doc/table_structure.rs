//! Representation-independent MS-DOC table grammar and grid planning.
//!
//! [MS-DOC] 2.4.3 encodes cells and rows with paragraph marks. This module is
//! the single owner of that nesting grammar and of the union-grid construction
//! used by both the legacy DOCX serializer and the direct model producer.

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

#[derive(Default)]
struct Pending<P> {
    rows: Vec<(Row, Vec<P>)>,
    cells: Vec<P>,
    cell: P,
}

pub(super) struct Assembler<P> {
    stack: Vec<Pending<P>>,
    body: P,
    rows: usize,
}

pub(super) struct Event<P>(pub(super) Vec<LogicalTable<P>>);

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

impl<P: Payload> Assembler<P> {
    pub(super) fn new() -> Self {
        Self {
            stack: Vec::new(),
            body: P::default(),
            rows: 0,
        }
    }

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
                || current.cells.len() != props.row.cells.len()
                || current.cells.is_empty()
            {
                return Err(unsupported("Word row definition does not match cell marks"));
            }
            reserve_one(&mut current.rows, admit)?;
            current
                .rows
                .push((props.row, std::mem::take(&mut current.cells)));
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

    fn close<F, A>(&mut self, emit: &mut F, admit: &mut A) -> Result<(), String>
    where
        F: FnMut(Event<P>) -> Result<P, String>,
        A: FnMut(usize) -> Result<(), String>,
    {
        let pending = self.stack.pop().expect("open table");
        if !pending.cell.is_empty() || !pending.cells.is_empty() {
            return Err(unsupported("unterminated Word table row"));
        }
        let mut tables = Vec::new();
        let mut rows = pending.rows.into_iter().peekable();
        while let Some(first) = rows.next() {
            let bidi = first.0.bidi;
            let mut group = Vec::new();
            reserve_one(&mut group, admit)?;
            group.push(first);
            while rows
                .peek()
                .is_some_and(|(row, _)| row.bidi == bidi && row.identity == group[0].0.identity)
            {
                reserve_one(&mut group, admit)?;
                group.push(rows.next().expect("peeked row"));
            }
            reserve_one(&mut tables, admit)?;
            tables.push(plan(group, admit)?);
        }
        let payload = emit(Event(tables))?;
        if let Some(parent) = self.stack.last_mut() {
            parent.cell.append(payload, admit)?;
        } else {
            self.body.append(payload, admit)?;
        }
        Ok(())
    }

    pub(super) fn finish<F, A>(mut self, mut emit: F, admit: &mut A) -> Result<P, String>
    where
        F: FnMut(Event<P>) -> Result<P, String>,
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

fn plan<P: Default, A: FnMut(usize) -> Result<(), String>>(
    rows: Vec<(Row, Vec<P>)>,
    admit: &mut A,
) -> Result<LogicalTable<P>, String> {
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
    for (mut row, contents) in rows {
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
    rows: &[(Row, Vec<P>)],
    admit: &mut A,
) -> Result<(Vec<(i32, usize)>, usize), String> {
    let edge_count = rows.iter().try_fold(0usize, |sum, (row, _)| {
        sum.checked_add(row.cells.len().checked_add(1).ok_or("OUTPUT_TOO_LARGE")?)
            .ok_or("OUTPUT_TOO_LARGE")
    })?;
    let mut boundaries = Vec::<(i32, usize)>::new();
    reserve_exact(&mut boundaries, edge_count, admit)?;
    for (row_index, (row, _)) in rows.iter().enumerate() {
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

    fn cell(width: i32) -> Cell {
        Cell {
            width,
            ..Cell::default()
        }
    }

    #[test]
    fn reverse_ordered_distinct_edges_sort_once_with_bounded_comparisons() {
        let row_count = 32_768usize;
        let mut rows = Vec::<(Row, Vec<()>)>::with_capacity(row_count);
        for index in 0..row_count {
            let mut row = Row::default();
            row.left = i32::try_from((row_count - index) * 2).unwrap();
            row.cells.push(cell(1));
            rows.push((row, Vec::new()));
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
            &[(first, Vec::<()>::new()), (second, Vec::new())],
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
            plan(vec![(row, vec![(); 65_536])], &mut |_| Ok(()))
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
        let error = grid_boundaries(&[(row, Vec::<()>::new())], &mut |bytes| {
            calls.push(bytes);
            Err("denied".into())
        })
        .unwrap_err();
        assert_eq!(error, "denied");
        assert_eq!(calls, vec![expected]);
    }
}
