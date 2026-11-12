//! The resource governor of the OpenDocument readers.
//!
//! Memory is measured, not estimated. [`Metered`] is the crate's global
//! allocator: it counts every allocation of the reading thread, so nothing a
//! reader keeps or copies can be left out of the account, and no reader
//! computes sizes. A reader runs inside a [`Scope`]; from then on
//!
//! * the bytes allocated beyond the scope's start never exceed the scope's
//!   limit in a document that is returned ([`Scope::finish`] checks the
//!   peak), and never exceed the limit plus [`CEILING_SLACK`] at any
//!   moment: an allocation past that hard ceiling fails, which aborts the
//!   read (a WASM trap) instead of growing memory. The slack is twice the
//!   largest package part a caller can allow, so single items copied
//!   between two checks stay under it; it stops a loop that allocates
//!   without checking;
//! * work is bounded by [`MAX_WORK`] steps, a step being one allocation
//!   (counted by the allocator) or one explicit [`tick`], and by
//!   [`MAX_COPIED_BYTES`] of allocation volume, which bounds copying of
//!   shared data however it is reached.
//!
//! The rules a reader follows are therefore two, and neither involves sizes:
//!
//! 1. every loop whose trip count comes from the file rather than from the
//!    XML nodes it visits (repeat counts, spans, cell ranges, `text:c`) runs
//!    through [`each`] / [`tick`], so repetition is walked item by item and
//!    never allocated ahead (no `vec![..; count]`, `with_capacity(count)` or
//!    `repeat(count)` on a file-derived count);
//! 2. walks that allocate per XML node tick per node, so the limits are
//!    checked while a large document is read, not after.
//!
//! A missed tick cannot leak memory past the hard ceiling and cannot hide an
//! over-limit document from [`Scope::finish`]; it can only delay the
//! rejection.
//!
//! What is bounded is the growth of the reading thread's heap since the
//! scope began: the quantity a host has to provide. Memory that existed
//! before the scope and is freed inside it is credited, since the heap did
//! not grow by what replaced it. A reader therefore enters its scope once
//! its inputs are loaded and only borrows them from then on (the package,
//! its XML text and parsed documents), so that the growth is the memory of
//! the read itself.
//!
//! # What the scope does not cover
//!
//! The scope governs building the model. The phases around it have bounds
//! of their own, each a function of the package limits the caller sets
//! (`PackageLimits`; the standard limits are 4,096 entries, 128 MiB per
//! entry and 256 MiB inflated in total, the hard ceilings 20,000 entries,
//! 512 MiB and 1 GiB):
//!
//! * the input: the file's bytes as the host hands them over, which the
//!   package owns for the session (the reader makes no copy of them);
//! * opening the package: the central directory, one record per entry, at
//!   most the entry limit;
//! * loading the XML parts (`styles.xml`, `content.xml`, `settings.xml`):
//!   their inflated text, each at most the entry limit and together at most
//!   the total inflated limit;
//! * parsing them: `xml::preflight` bounds each part before the parser
//!   runs, to [`crate::xml::MAX_NODES_PER_PART`] nodes,
//!   [`crate::xml::MAX_XML_RESERVATION_BYTES`] of parser records and
//!   [`crate::xml::MAX_XML_WORK`] units of parsing work;
//! * handing the model to the host (the spreadsheet session): the workbook
//!   bootstrap and each worksheet chunk are written through bounded
//!   writers (the workbook and worksheet JSON limits of the shared
//!   resource policy), rows are lent at most 128 at a time, and an image
//!   is one package entry, at most the standard entry limit.
//!
//! Chart objects and images read while the model is built are read inside
//! the scope.

use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

/// Implementation bound on what one document's reader holds on the heap
/// beyond its parsed XML: the model it returns and everything transient.
pub const MAX_MODEL_BYTES: usize = 512 * 1024 * 1024;
/// Implementation bound on the steps of one read (allocations plus explicit
/// ticks).
pub const MAX_WORK: u64 = 128 * 1024 * 1024;
/// Implementation bound on the bytes one read allocates in total (retained
/// or not): the copying it may do.
pub const MAX_COPIED_BYTES: u64 = 8 * MAX_MODEL_BYTES as u64;
/// How far past its limit a scope may be between two checks before an
/// allocation fails: twice the hard bound on one inflated package part
/// (the shared resource policy's `HARD_MAX_ARCHIVE_ENTRY_BYTES`, 512 MiB).
pub const CEILING_SLACK: usize = 2 * MAX_MODEL_BYTES;

thread_local! {
    /// Bytes currently allocated by this thread (frees on another thread
    /// are not seen; the readers are single-threaded).
    static LIVE: Cell<i64> = const { Cell::new(0) };
    static PEAK: Cell<i64> = const { Cell::new(0) };
    static ALLOCATIONS: Cell<u64> = const { Cell::new(0) };
    static ALLOCATED: Cell<u64> = const { Cell::new(0) };
    /// The value of [`LIVE`] no allocation may pass.
    static CEILING: Cell<i64> = const { Cell::new(i64::MAX) };
    static CURRENT: Cell<Option<State>> = const { Cell::new(None) };
}

/// The counting global allocator (over the system allocator).
pub struct Metered;

/// Account `grow` more live bytes of an allocation of `size` bytes; `false`
/// when it would pass the hard ceiling.
fn admit(grow: i64, size: usize) -> bool {
    LIVE.try_with(|live| {
        let next = live.get().saturating_add(grow);
        if grow > 0 && next > CEILING.with(Cell::get) {
            return false;
        }
        live.set(next);
        PEAK.with(|peak| peak.set(peak.get().max(next)));
        ALLOCATIONS.with(|count| count.set(count.get().saturating_add(1)));
        ALLOCATED.with(|bytes| bytes.set(bytes.get().saturating_add(size as u64)));
        true
    })
    // Thread teardown: nothing is being read.
    .unwrap_or(true)
}

fn release(size: usize) {
    let _ = LIVE.try_with(|live| live.set(live.get().saturating_sub(size as i64)));
}

// SAFETY: every request is forwarded unchanged to `System`; the counters are
// thread-local cells without destructors, so reading them never allocates.
unsafe impl GlobalAlloc for Metered {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        if !admit(layout.size() as i64, layout.size()) {
            return std::ptr::null_mut();
        }
        System.alloc(layout)
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        if !admit(layout.size() as i64, layout.size()) {
            return std::ptr::null_mut();
        }
        System.alloc_zeroed(layout)
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        release(layout.size());
        System.dealloc(ptr, layout)
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        if !admit(new_size as i64 - layout.size() as i64, new_size) {
            return std::ptr::null_mut();
        }
        System.realloc(ptr, layout, new_size)
    }
}

#[derive(Clone, Copy)]
struct State {
    /// [`LIVE`] when the scope began.
    base: i64,
    limit: i64,
    allocations: u64,
    allocated: u64,
    ticks: u64,
    cells: u64,
    max_cells: u64,
}

/// One governed read on this thread. Dropping it ends the governance.
pub struct Scope {
    /// `None` when this scope joined one already running.
    outer: Option<(i64, i64)>,
    /// Scopes are bound to their thread's counters.
    _thread: std::marker::PhantomData<*const ()>,
}

impl Scope {
    /// Govern the read that follows with `limit` bytes (never more than
    /// [`MAX_MODEL_BYTES`]). Inside a running scope this joins it: a chart
    /// read for a workbook is part of the workbook's read.
    pub fn enter(limit: usize) -> Self {
        if CURRENT.with(Cell::get).is_some() {
            return Self {
                outer: None,
                _thread: std::marker::PhantomData,
            };
        }
        let limit = limit.min(MAX_MODEL_BYTES) as i64;
        let base = LIVE.with(Cell::get);
        let outer = (CEILING.with(Cell::get), PEAK.with(Cell::get));
        PEAK.with(|peak| peak.set(base));
        CEILING.with(|ceiling| {
            ceiling.set(
                base.saturating_add(limit)
                    .saturating_add(CEILING_SLACK as i64),
            )
        });
        CURRENT.with(|current| {
            current.set(Some(State {
                base,
                limit,
                allocations: ALLOCATIONS.with(Cell::get),
                allocated: ALLOCATED.with(Cell::get),
                ticks: 0,
                cells: 0,
                max_cells: u64::MAX,
            }))
        });
        Self {
            outer: Some(outer),
            _thread: std::marker::PhantomData,
        }
    }

    /// Bound the items [`cell`] admits in this scope.
    pub fn with_cell_limit(self, max_cells: usize) -> Self {
        CURRENT.with(|current| {
            if let Some(mut state) = current.get() {
                state.max_cells = max_cells as u64;
                current.set(Some(state));
            }
        });
        self
    }

    /// The final check: `Err` when the read ever passed a limit, including
    /// between two ticks.
    pub fn finish(self) -> Result<(), String> {
        tick(0)
    }
}

impl Drop for Scope {
    fn drop(&mut self) {
        if let Some((ceiling, peak)) = self.outer {
            CURRENT.with(|current| current.set(None));
            CEILING.with(|cell| cell.set(ceiling));
            PEAK.with(|cell| cell.set(cell.get().max(peak)));
        }
    }
}

/// A resource-limit rejection (content-free).
fn over(what: &str) -> String {
    crate::error::resource(what)
}

/// Count `steps` of work and check every limit of the running scope. Outside
/// a scope (helpers called on their own) nothing is limited.
pub fn tick(steps: usize) -> Result<(), String> {
    CURRENT.with(|current| {
        let Some(mut state) = current.get() else {
            return Ok(());
        };
        state.ticks = state.ticks.saturating_add(steps as u64);
        current.set(Some(state));
        if PEAK.with(Cell::get).saturating_sub(state.base) > state.limit {
            return Err(over("model size"));
        }
        let allocations = ALLOCATIONS.with(Cell::get) - state.allocations;
        if state.ticks.saturating_add(allocations) > MAX_WORK {
            return Err(over("reading work"));
        }
        if ALLOCATED.with(Cell::get) - state.allocated > MAX_COPIED_BYTES {
            return Err(over("copied data"));
        }
        Ok(())
    })
}

/// Admit one more materialized item (a cell or a row record) of the scope's
/// item bound, and one step.
pub fn cell() -> Result<(), String> {
    CURRENT.with(|current| {
        if let Some(mut state) = current.get() {
            state.cells = state.cells.saturating_add(1);
            current.set(Some(state));
            if state.cells > state.max_cells {
                return Err(over("cell count"));
            }
        }
        Ok(())
    })?;
    tick(1)
}

/// `first..=last`, one step per item: the way to walk a repetition or a
/// range whose length the file chose.
pub fn each(first: u32, last: u32) -> impl Iterator<Item = Result<u32, String>> {
    let mut next = Some(first).filter(|first| *first <= last);
    let mut failed = false;
    std::iter::from_fn(move || {
        if failed {
            return None;
        }
        let item = next?;
        next = item.checked_add(1).filter(|n| *n <= last);
        Some(tick(1).map(|()| item).inspect_err(|_| failed = true))
    })
}

/// Append `count` spaces (a `text:s` run, §6.1.3) in bounded pieces, a step
/// each, so a huge count is rejected while it is written, never allocated
/// at once.
pub fn push_spaces(out: &mut String, count: usize) -> Result<(), String> {
    const PIECE: usize = 4096;
    let mut left = count;
    while left > 0 {
        let piece = left.min(PIECE);
        tick(1)?;
        out.extend(std::iter::repeat_n(' ', piece));
        left -= piece;
    }
    Ok(())
}

/// Bytes the running scope held at its peak (tests and diagnostics).
pub fn peak_bytes() -> usize {
    CURRENT.with(|current| {
        current.get().map_or(0, |state| {
            usize::try_from(PEAK.with(Cell::get).saturating_sub(state.base)).unwrap_or(0)
        })
    })
}

/// Steps the running scope has taken (tests and diagnostics).
pub fn work() -> u64 {
    CURRENT.with(|current| {
        current.get().map_or(0, |state| {
            state
                .ticks
                .saturating_add(ALLOCATIONS.with(Cell::get) - state.allocations)
        })
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_allocation_past_the_hard_ceiling_is_refused() {
        let _scope = Scope::enter(1 << 20);
        let within = (1 << 20) + CEILING_SLACK as i64;
        assert!(admit(within, 0));
        release(within as usize);
        assert!(!admit(within + 1, 0));
        // A refused allocation is not counted as live.
        assert!(admit(within, 0));
        release(within as usize);
    }

    #[test]
    fn the_allocator_meters_what_a_scope_holds() {
        let scope = Scope::enter(1 << 20);
        let held = vec![0u8; 300_000];
        assert!(peak_bytes() >= 300_000);
        assert!(tick(1).is_ok());
        let more = vec![0u8; 900_000];
        // Over the limit: every later check rejects, even once it is freed.
        assert!(tick(1).is_err());
        drop(more);
        drop(held);
        assert!(scope.finish().is_err());
        // The next scope starts clean.
        let scope = Scope::enter(1 << 20);
        assert!(scope.finish().is_ok());
    }

    #[test]
    fn walks_and_cells_are_bounded() {
        let _scope = Scope::enter(1 << 20).with_cell_limit(2);
        assert_eq!(each(3, 5).collect::<Result<Vec<_>, _>>(), Ok(vec![3, 4, 5]));
        assert_eq!(each(u32::MAX, u32::MAX).count(), 1);
        assert_eq!(each(2, 1).count(), 0);
        assert!(cell().is_ok() && cell().is_ok());
        assert!(cell().is_err());
        assert!(tick(MAX_WORK as usize).is_err());
    }

    #[test]
    fn spaces_are_written_in_pieces() {
        let _scope = Scope::enter(1 << 20);
        let mut out = String::new();
        push_spaces(&mut out, 10_000).unwrap();
        assert_eq!(out.len(), 10_000);
        assert!(push_spaces(&mut out, usize::MAX).is_err());
        assert!(out.len() < 4 << 20);
    }
}
