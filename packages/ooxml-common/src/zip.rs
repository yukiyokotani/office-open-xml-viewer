//! Shared configuration for ZIP entry decompression limits.
//!
//! OOXML parsers must cap per-entry decompressed output to block zip-bomb DoS.
//! The cap defaults to 512 MiB — large enough for legitimate embedded video /
//! 4K media but small enough to refuse pathological archives — and can be
//! overridden per-parse via the `wasm_bindgen` entry points so library users
//! can tighten the budget (untrusted gateways) or loosen it (legitimate huge
//! decks) without forking.
//!
//! The current cap is stored in a thread-local so existing `read_zip_*`
//! helpers in each parser can consult it without threading a parameter
//! through ~70 call sites. Each WASM entry point installs a [`Guard`] for
//! its scope and the cap is restored on drop, so concurrent JS callers never
//! interfere (WASM is single-threaded; each invocation runs to completion).

use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::rc::Rc;

/// 512 MiB. OOXML legitimately reaches tens of MB (embedded video, 4K
/// images) but not hundreds, so this cap blocks zip-bomb DoS without
/// rejecting real files.
pub const DEFAULT_MAX_ZIP_ENTRY_BYTES: u64 = 512 * 1024 * 1024;

/// 512 MiB across all distinct entries actually inflated from one archive.
/// This bounds archives made from many individually-valid entries while keeping
/// ordinary OOXML packages (whose XML and media total far below this) compatible.
pub const DEFAULT_MAX_ZIP_TOTAL_BYTES: u64 = 512 * 1024 * 1024;

/// A conservative ceiling for central-directory entries. OOXML packages normally
/// contain hundreds or a few thousand parts; 20,000 leaves ample headroom while
/// bounding metadata scans and pathological part explosions.
pub const DEFAULT_MAX_ZIP_ENTRIES: u64 = 20_000;

/// Upper bound on the buffer we pre-reserve from an entry's DECLARED size.
///
/// `entry.size()` is the uncompressed size recorded in the zip header, which is
/// attacker-controlled: a zip-bomb variant can declare 512 MiB (up to the entry
/// cap) while the real payload is a few bytes. Feeding that straight into
/// `Vec::with_capacity` wastes up to `DEFAULT_MAX_ZIP_ENTRY_BYTES` of eager
/// allocation per entry. We instead pre-reserve at most 1 MiB and let
/// `read_to_end` grow the buffer for genuinely large parts.
///
/// 1 MiB is chosen because it comfortably fits the vast majority of OOXML parts
/// (document.xml / sheetN.xml / slideN.xml are typically tens to a few hundred
/// KiB) so they incur zero reallocation, while capping the wasted reserve for a
/// forged header at 1 MiB. Genuinely large parts (multi-MB sharedStrings) grow
/// via `read_to_end`'s amortized-O(n) doubling — a handful of reallocs, no
/// measurable parse-time cost (verified against the parse bench).
const INITIAL_RESERVE_CAP: usize = 1024 * 1024; // 1 MiB

/// Buffer capacity to pre-reserve for an entry that declares `declared_size`
/// uncompressed bytes: the declared size when small, else [`INITIAL_RESERVE_CAP`].
/// Clamps the eager `with_capacity` so a forged declaration cannot force a giant
/// up-front allocation; the read still completes in full via `read_to_end`.
fn initial_reserve(declared_size: u64) -> usize {
    declared_size.min(INITIAL_RESERVE_CAP as u64) as usize
}

thread_local! {
    static MAX_ZIP_ENTRY_BYTES: Cell<u64> = const { Cell::new(DEFAULT_MAX_ZIP_ENTRY_BYTES) };
    static ACTIVE_BUDGET: RefCell<Option<ZipBudget>> = const { RefCell::new(None) };
}

#[derive(Clone)]
pub struct ZipBudget(Rc<RefCell<BudgetState>>);

struct BudgetState {
    max_entry_bytes: u64,
    max_total_bytes: u64,
    max_entries: u64,
    /// The greatest number of bytes observed for each entry. Re-reading an entry
    /// through a retained archive replaces this value instead of double-counting.
    entry_bytes: HashMap<String, u64>,
    actual_total_bytes: u64,
    /// A retained archive fails closed after a proven limit breach. Without this,
    /// callers could catch one error and repeatedly inflate more distinct forged
    /// entries through the same handle.
    poisoned: bool,
}

impl ZipBudget {
    pub fn new(
        max_zip_entry_bytes: Option<u64>,
        max_zip_total_bytes: Option<u64>,
        max_zip_entries: Option<u64>,
    ) -> Self {
        Self(Rc::new(RefCell::new(BudgetState {
            max_entry_bytes: max_zip_entry_bytes
                .filter(|&n| n > 0)
                .unwrap_or(DEFAULT_MAX_ZIP_ENTRY_BYTES),
            max_total_bytes: max_zip_total_bytes
                .filter(|&n| n > 0)
                .unwrap_or(DEFAULT_MAX_ZIP_TOTAL_BYTES),
            max_entries: max_zip_entries
                .filter(|&n| n > 0)
                .unwrap_or(DEFAULT_MAX_ZIP_ENTRIES),
            entry_bytes: HashMap::new(),
            actual_total_bytes: 0,
            poisoned: false,
        })))
    }
}

/// RAII guard that restores the previous cap when dropped. Created by
/// [`scoped_max`]; the caller should bind it to a `let _guard = …` for the
/// full duration of the parse call.
#[must_use = "binding the guard keeps the cap installed for this scope"]
pub struct Guard {
    previous: u64,
    previous_budget: Option<ZipBudget>,
}

impl Drop for Guard {
    fn drop(&mut self) {
        MAX_ZIP_ENTRY_BYTES.with(|c| c.set(self.previous));
        ACTIVE_BUDGET.with(|budget| *budget.borrow_mut() = self.previous_budget.take());
    }
}

/// Install a per-call ZIP entry size cap for the lifetime of the returned
/// guard. `None`, zero, or any non-positive value falls back to
/// [`DEFAULT_MAX_ZIP_ENTRY_BYTES`].
pub fn scoped_max(value: Option<u64>) -> Guard {
    scoped_limits(value, None, None)
}

/// Create and install a ZIP budget for one non-retained parse or extraction.
pub fn scoped_limits(
    max_zip_entry_bytes: Option<u64>,
    max_zip_total_bytes: Option<u64>,
    max_zip_entries: Option<u64>,
) -> Guard {
    let budget = ZipBudget::new(max_zip_entry_bytes, max_zip_total_bytes, max_zip_entries);
    scoped_budget(&budget)
}

/// Install a retained archive's budget for the duration of one method call.
/// The handle owns this budget, so actual bytes remain accounted across calls.
pub fn scoped_budget(budget: &ZipBudget) -> Guard {
    let resolved = budget.0.borrow().max_entry_bytes;
    let previous = MAX_ZIP_ENTRY_BYTES.with(|c| c.replace(resolved));
    let previous_budget = ACTIVE_BUDGET.with(|active| active.borrow_mut().replace(budget.clone()));
    Guard {
        previous,
        previous_budget,
    }
}

/// Current cap in effect on this thread. Parsers consult this from their
/// `read_zip_*` helpers when validating entry sizes.
pub fn current_max() -> u64 {
    MAX_ZIP_ENTRY_BYTES.with(Cell::get)
}

fn current_limits() -> (u64, u64, u64) {
    ACTIVE_BUDGET.with(|active| {
        active
            .borrow()
            .as_ref()
            .map(|budget| {
                let state = budget.0.borrow();
                (
                    state.max_entry_bytes,
                    state.max_total_bytes,
                    state.max_entries,
                )
            })
            .unwrap_or((
                DEFAULT_MAX_ZIP_ENTRY_BYTES,
                DEFAULT_MAX_ZIP_TOTAL_BYTES,
                DEFAULT_MAX_ZIP_ENTRIES,
            ))
    })
}

fn ensure_budget_is_healthy() -> Result<(), String> {
    ACTIVE_BUDGET.with(|active| {
        let Some(budget) = active.borrow().clone() else {
            return Ok(());
        };
        if budget.0.borrow().poisoned {
            return Err("ZIP archive budget already exceeded".to_string());
        }
        Ok(())
    })
}

fn poison_budget(message: &'static str) -> String {
    ACTIVE_BUDGET.with(|active| {
        if let Some(budget) = active.borrow().clone() {
            budget.0.borrow_mut().poisoned = true;
        }
    });
    message.to_string()
}

/// Reject an archive whose central directory already proves it exceeds a limit.
/// The declared sizes are not trusted as actual bytes, but they are a safe early
/// rejection signal that avoids inflating or allocating for a known-over-budget
/// archive. Actual decompressed bytes are checked separately while reading.
pub fn validate_archive_limits<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Result<(), String> {
    ensure_budget_is_healthy()?;
    let (_, max_total, max_entries) = current_limits();
    if (archive.len() as u64) > max_entries {
        return Err(poison_budget("ZIP archive exceeds entry count limit"));
    }
    let mut declared_total = 0u64;
    for index in 0..archive.len() {
        let entry = archive
            .by_index(index)
            .map_err(|_| "ZIP archive entry metadata error".to_string())?;
        declared_total = declared_total
            .checked_add(entry.size())
            .ok_or_else(|| poison_budget("ZIP archive exceeds total size limit"))?;
        if declared_total > max_total {
            return Err(poison_budget("ZIP archive exceeds total size limit"));
        }
    }
    Ok(())
}

fn observe_actual_bytes(path: &str, observed_for_entry: u64) -> Result<(), String> {
    ensure_budget_is_healthy()?;
    ACTIVE_BUDGET.with(|active| {
        let Some(budget) = active.borrow().clone() else {
            return Ok(());
        };
        let mut state = budget.0.borrow_mut();
        let old = state.entry_bytes.get(path).copied().unwrap_or(0);
        if observed_for_entry <= old {
            return Ok(());
        }
        let next_total = state
            .actual_total_bytes
            .checked_add(observed_for_entry - old)
            .ok_or_else(|| {
                state.poisoned = true;
                "ZIP archive exceeds total size limit".to_string()
            })?;
        if next_total > state.max_total_bytes {
            state.poisoned = true;
            return Err("ZIP archive exceeds total size limit".to_string());
        }
        state
            .entry_bytes
            .insert(path.to_owned(), observed_for_entry);
        state.actual_total_bytes = next_total;
        Ok(())
    })
}

fn read_limited_bytes(
    reader: &mut impl std::io::Read,
    path: &str,
    declared_size: u64,
    max_bytes: Option<u64>,
) -> Result<Vec<u8>, String> {
    ensure_budget_is_healthy()?;
    let (max_entry, _, _) = current_limits();
    if declared_size > max_entry {
        return Err(poison_budget("ZIP entry exceeds size limit"));
    }
    let limit = max_bytes.unwrap_or(max_entry).min(max_entry);
    let mut buf = Vec::with_capacity(initial_reserve(declared_size.min(limit)));
    let mut chunk = [0u8; 32 * 1024];
    let mut read = 0u64;
    loop {
        if read == limit {
            // Prefix probes intentionally stop here. Full reads consume one more
            // byte so an entry whose actual inflate output exceeds its cap cannot
            // masquerade as an exact-limit entry behind forged metadata.
            if max_bytes.is_some() {
                break;
            }
            let count = reader
                .read(&mut chunk[..1])
                .map_err(|_| "ZIP entry read error".to_string())?;
            if count != 0 {
                return Err(poison_budget("ZIP entry exceeds size limit"));
            }
            break;
        }
        let remaining = (limit - read).min(chunk.len() as u64) as usize;
        let count = reader
            .read(&mut chunk[..remaining])
            .map_err(|_| "ZIP entry read error".to_string())?;
        if count == 0 {
            break;
        }
        read = read
            .checked_add(count as u64)
            .ok_or_else(|| poison_budget("ZIP entry exceeds size limit"))?;
        if read > limit {
            return Err(poison_budget("ZIP entry exceeds size limit"));
        }
        observe_actual_bytes(path, read)?;
        buf.extend_from_slice(&chunk[..count]);
    }
    Ok(buf)
}

/// Read one zip entry's bytes by path. Honors the scoped max-entry guard:
/// entries whose declared size exceeds the cap (default 512 MiB, or the
/// per-call override) are rejected rather than truncated — the zip-bomb DoS
/// guard shared with the per-parser `extract_*` WASM entry points.
pub fn extract_zip_entry(
    data: &[u8],
    path: &str,
    max_zip_entry_bytes: Option<u64>,
) -> Result<Vec<u8>, String> {
    extract_zip_entry_with_limits(data, path, max_zip_entry_bytes, None, None)
}

/// Limit-aware counterpart of [`extract_zip_entry`]. Browser entry points use
/// this to apply entry, aggregate-byte, and central-directory-count limits.
pub fn extract_zip_entry_with_limits(
    data: &[u8],
    path: &str,
    max_zip_entry_bytes: Option<u64>,
    max_zip_total_bytes: Option<u64>,
    max_zip_entries: Option<u64>,
) -> Result<Vec<u8>, String> {
    use std::io::Cursor;
    let _guard = scoped_limits(max_zip_entry_bytes, max_zip_total_bytes, max_zip_entries);
    let cursor = Cursor::new(data);
    let mut zip = zip::ZipArchive::new(cursor).map_err(|e| format!("zip open error: {e}"))?;
    validate_archive_limits(&mut zip)?;
    let mut entry = zip
        .by_name(path)
        .map_err(|e| format!("entry not found: {path}: {e}"))?;
    read_limited_bytes(&mut entry, path, entry.size(), None)
}

/// Read one entry's bytes from an **already-opened** [`ZipArchive`]. Twin of
/// [`extract_zip_entry`] for callers that keep a single archive open across
/// many reads (the common case inside a parser) instead of re-opening it from
/// the raw bytes per entry. Honors the scoped max-entry guard: an entry whose
/// declared size exceeds the current cap is rejected with an `Err`, never
/// silently truncated (the zip-bomb DoS guard). Generic over the archive's
/// reader so each parser's concrete type (`Cursor<&[u8]>`, …) works unchanged.
pub fn read_zip_bytes<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<Vec<u8>, String> {
    ensure_budget_is_healthy()?;
    let mut entry = archive
        .by_name(path)
        .map_err(|e| format!("entry not found: {path}: {e}"))?;
    read_limited_bytes(&mut entry, path, entry.size(), None)
}

/// UTF-8 string counterpart of [`read_zip_bytes`] for XML parts. Same cap
/// enforcement and archive-reuse contract; decodes the entry as UTF-8 (strict —
/// OOXML parts are well-formed UTF-8, and a decode failure is a real corruption
/// worth reporting rather than papering over with lossy substitution).
pub fn read_zip_string<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<String, String> {
    let bytes = read_zip_bytes(archive, path)?;
    String::from_utf8(bytes).map_err(|_| "ZIP entry is not valid UTF-8".to_string())
}

/// Read a bounded prefix of an entry while accounting actual bytes in the same
/// retained-archive budget. A later full read of the same entry raises its
/// observed total to the full size rather than charging the prefix twice.
pub fn read_zip_prefix_bytes<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
    max_bytes: u64,
) -> Result<Vec<u8>, String> {
    ensure_budget_is_healthy()?;
    let mut entry = archive
        .by_name(path)
        .map_err(|e| format!("entry not found: {path}: {e}"))?;
    read_limited_bytes(&mut entry, path, entry.size(), Some(max_bytes))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn aggregate_budget_rejects_actual_bytes_across_distinct_entries() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default();
            w.start_file("one.bin", opts).unwrap();
            w.write_all(b"1234").unwrap();
            w.start_file("two.bin", opts).unwrap();
            w.write_all(b"5678").unwrap();
            w.finish().unwrap();
        }
        let mut archive = zip::ZipArchive::new(Cursor::new(buf)).unwrap();
        let _guard = scoped_limits(Some(16), Some(6), Some(2));
        assert_eq!(read_zip_bytes(&mut archive, "one.bin").unwrap(), b"1234");
        let err = read_zip_bytes(&mut archive, "two.bin").unwrap_err();
        assert!(err.contains("total size limit"), "got: {err}");
    }

    #[test]
    fn retained_budget_does_not_double_charge_an_entry_read_again() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default();
            w.start_file("one.bin", opts).unwrap();
            w.write_all(b"1234").unwrap();
            w.start_file("two.bin", opts).unwrap();
            w.write_all(b"5678").unwrap();
            w.finish().unwrap();
        }
        let mut archive = zip::ZipArchive::new(Cursor::new(buf)).unwrap();
        let budget = ZipBudget::new(Some(16), Some(8), Some(2));
        {
            let _guard = scoped_budget(&budget);
            assert_eq!(read_zip_bytes(&mut archive, "one.bin").unwrap(), b"1234");
        }
        {
            let _guard = scoped_budget(&budget);
            assert_eq!(read_zip_bytes(&mut archive, "one.bin").unwrap(), b"1234");
            assert_eq!(read_zip_bytes(&mut archive, "two.bin").unwrap(), b"5678");
        }
    }

    #[test]
    fn retained_budget_stays_rejected_after_aggregate_crossing() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default();
            w.start_file("one.bin", opts).unwrap();
            w.write_all(b"1234").unwrap();
            w.start_file("two.bin", opts).unwrap();
            w.write_all(b"5678").unwrap();
            w.finish().unwrap();
        }
        let mut archive = zip::ZipArchive::new(Cursor::new(buf)).unwrap();
        let budget = ZipBudget::new(Some(16), Some(6), Some(2));
        {
            let _guard = scoped_budget(&budget);
            assert_eq!(read_zip_bytes(&mut archive, "one.bin").unwrap(), b"1234");
            let err = read_zip_bytes(&mut archive, "two.bin").unwrap_err();
            assert!(err.contains("total size limit"), "got: {err}");
        }
        {
            let _guard = scoped_budget(&budget);
            let err = read_zip_bytes(&mut archive, "one.bin").unwrap_err();
            assert!(err.contains("budget already exceeded"), "got: {err}");
        }
    }

    #[test]
    fn retained_budget_stays_rejected_after_actual_entry_overflow() {
        use std::io::Cursor;
        // A forged stream can emit more bytes than its claimed uncompressed
        // size. Exercise the shared reader directly so the test does not rely on
        // a ZIP implementation accepting malformed central-directory metadata.
        let budget = ZipBudget::new(Some(4), Some(16), Some(2));
        {
            let _guard = scoped_budget(&budget);
            let mut forged = Cursor::new(b"12345");
            let err = read_limited_bytes(&mut forged, "forged.bin", 4, None).unwrap_err();
            assert!(err.contains("entry exceeds size limit"), "got: {err}");
        }
        {
            let _guard = scoped_budget(&budget);
            let mut harmless = Cursor::new(b"1");
            let err = read_limited_bytes(&mut harmless, "later.bin", 1, None).unwrap_err();
            assert!(err.contains("budget already exceeded"), "got: {err}");
        }
    }

    #[test]
    fn central_directory_entry_count_is_rejected_before_reads() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default();
            w.start_file("one.xml", opts).unwrap();
            w.write_all(b"1").unwrap();
            w.start_file("two.xml", opts).unwrap();
            w.write_all(b"2").unwrap();
            w.finish().unwrap();
        }
        let mut archive = zip::ZipArchive::new(Cursor::new(buf)).unwrap();
        let _guard = scoped_limits(Some(16), Some(16), Some(1));
        let err = validate_archive_limits(&mut archive).unwrap_err();
        assert!(err.contains("entry count limit"), "got: {err}");
    }

    #[test]
    fn extract_zip_entry_reads_by_path() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default();
            w.start_file("ppt/media/image1.png", opts).unwrap();
            w.write_all(b"\x89PNGdata").unwrap();
            w.finish().unwrap();
        }
        let bytes = extract_zip_entry(&buf, "ppt/media/image1.png", None).unwrap();
        assert_eq!(bytes, b"\x89PNGdata");
        assert!(extract_zip_entry(&buf, "ppt/media/missing.png", None)
            .unwrap_err()
            .contains("not found"));
    }

    #[test]
    fn extract_zip_entry_rejects_oversized_entry() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default();
            w.start_file("ppt/media/big.bin", opts).unwrap();
            w.write_all(b"12345678").unwrap(); // 8 bytes uncompressed
            w.finish().unwrap();
        }
        // A cap below the declared size must be REJECTED, never silently
        // truncated — this is the zip-bomb DoS guard (default 512 MiB).
        let err = extract_zip_entry(&buf, "ppt/media/big.bin", Some(4)).unwrap_err();
        assert!(err.contains("exceeds size limit"), "got: {err}");
        // A cap above the size reads the entry in full.
        assert_eq!(
            extract_zip_entry(&buf, "ppt/media/big.bin", Some(64)).unwrap(),
            b"12345678"
        );
    }

    /// Build a one-entry in-memory zip for the open-archive helper tests.
    fn archive_with(name: &str, body: &[u8]) -> zip::ZipArchive<std::io::Cursor<Vec<u8>>> {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default();
            w.start_file(name, opts).unwrap();
            w.write_all(body).unwrap();
            w.finish().unwrap();
        }
        zip::ZipArchive::new(Cursor::new(buf)).unwrap()
    }

    #[test]
    fn read_zip_bytes_reads_present_and_reports_missing() {
        let mut ar = archive_with("word/document.xml", b"<xml/>");
        assert_eq!(
            read_zip_bytes(&mut ar, "word/document.xml").unwrap(),
            b"<xml/>"
        );
        let err = read_zip_bytes(&mut ar, "word/missing.xml").unwrap_err();
        assert!(err.contains("not found"), "got: {err}");
    }

    #[test]
    fn read_zip_string_reads_present_and_reports_missing() {
        let mut ar = archive_with("xl/workbook.xml", b"<workbook/>");
        assert_eq!(
            read_zip_string(&mut ar, "xl/workbook.xml").unwrap(),
            "<workbook/>"
        );
        let err = read_zip_string(&mut ar, "xl/nope.xml").unwrap_err();
        assert!(err.contains("not found"), "got: {err}");
    }

    /// Forge a STORED (compression=0) single-entry zip whose declared
    /// `uncompressed_size` (in BOTH the local file header and the central
    /// directory) is much larger than the real body. Returns the raw zip bytes.
    ///
    /// A stored entry lets us set uncompressed==compressed==declared cleanly.
    /// We hand-lay the bytes so we control the size fields the way a malicious
    /// archive would. `real` is the actual payload; `declared` is the lie.
    #[cfg(test)]
    fn forged_stored_zip(name: &str, real: &[u8], declared: u32) -> Vec<u8> {
        let crc = {
            // CRC-32 of the real body (zip stores the checksum of actual data).
            const POLY: u32 = 0xEDB8_8320;
            let mut c: u32 = 0xFFFF_FFFF;
            for &b in real {
                c ^= b as u32;
                for _ in 0..8 {
                    c = if c & 1 != 0 { (c >> 1) ^ POLY } else { c >> 1 };
                }
            }
            !c
        };
        let name_bytes = name.as_bytes();
        let nlen = name_bytes.len() as u16;
        let mut z = Vec::new();
        // ── Local file header ──
        z.extend_from_slice(&0x0403_4b50u32.to_le_bytes()); // signature PK\x03\x04
        z.extend_from_slice(&20u16.to_le_bytes()); // version needed
        z.extend_from_slice(&0u16.to_le_bytes()); // flags
        z.extend_from_slice(&0u16.to_le_bytes()); // method = 0 (stored)
        z.extend_from_slice(&0u16.to_le_bytes()); // mod time
        z.extend_from_slice(&0u16.to_le_bytes()); // mod date
        z.extend_from_slice(&crc.to_le_bytes()); // crc-32
        z.extend_from_slice(&declared.to_le_bytes()); // compressed size (LIE)
        z.extend_from_slice(&declared.to_le_bytes()); // uncompressed size (LIE)
        z.extend_from_slice(&nlen.to_le_bytes()); // file name length
        z.extend_from_slice(&0u16.to_le_bytes()); // extra length
        z.extend_from_slice(name_bytes);
        let data_start = z.len();
        z.extend_from_slice(real); // only `real.len()` bytes actually present
                                   // ── Central directory header ──
        let cd_start = z.len();
        z.extend_from_slice(&0x0201_4b50u32.to_le_bytes()); // signature PK\x01\x02
        z.extend_from_slice(&20u16.to_le_bytes()); // version made by
        z.extend_from_slice(&20u16.to_le_bytes()); // version needed
        z.extend_from_slice(&0u16.to_le_bytes()); // flags
        z.extend_from_slice(&0u16.to_le_bytes()); // method = 0 (stored)
        z.extend_from_slice(&0u16.to_le_bytes()); // mod time
        z.extend_from_slice(&0u16.to_le_bytes()); // mod date
        z.extend_from_slice(&crc.to_le_bytes()); // crc-32
        z.extend_from_slice(&declared.to_le_bytes()); // compressed size (LIE)
        z.extend_from_slice(&declared.to_le_bytes()); // uncompressed size (LIE)
        z.extend_from_slice(&nlen.to_le_bytes()); // file name length
        z.extend_from_slice(&0u16.to_le_bytes()); // extra length
        z.extend_from_slice(&0u16.to_le_bytes()); // comment length
        z.extend_from_slice(&0u16.to_le_bytes()); // disk number start
        z.extend_from_slice(&0u16.to_le_bytes()); // internal attrs
        z.extend_from_slice(&0u32.to_le_bytes()); // external attrs
        z.extend_from_slice(&(data_start as u32 - 30 - nlen as u32).to_le_bytes()); // local header offset (=0)
        z.extend_from_slice(name_bytes);
        let cd_size = z.len() - cd_start;
        // ── End of central directory ──
        z.extend_from_slice(&0x0605_4b50u32.to_le_bytes()); // signature PK\x05\x06
        z.extend_from_slice(&0u16.to_le_bytes()); // disk number
        z.extend_from_slice(&0u16.to_le_bytes()); // cd start disk
        z.extend_from_slice(&1u16.to_le_bytes()); // entries on this disk
        z.extend_from_slice(&1u16.to_le_bytes()); // total entries
        z.extend_from_slice(&(cd_size as u32).to_le_bytes()); // cd size
        z.extend_from_slice(&(cd_start as u32).to_le_bytes()); // cd offset
        z.extend_from_slice(&0u16.to_le_bytes()); // comment length
        z
    }

    /// EMPIRICAL (RB11 attack-vector confirmation): `entry.size()` reports the
    /// DECLARED (attacker-controlled, central-directory) uncompressed size, NOT
    /// the actual decompressed byte count. This is the number the pre-fix code
    /// fed straight into `Vec::with_capacity`, so a forged header declaring
    /// 512 MiB over-reserved 512 MiB before reading a single byte. This test
    /// pins the observed behavior so a future zip-crate upgrade that changes it
    /// (returning the actual size, which would neutralize the vector) fails
    /// loudly.
    #[test]
    fn entry_size_reports_declared_not_actual() {
        use std::io::Cursor;
        // Declare 64 MiB but supply only 8 real bytes.
        const DECLARED: u32 = 64 * 1024 * 1024;
        let real = b"realdata"; // 8 bytes
        let raw = forged_stored_zip("word/document.xml", real, DECLARED);
        let mut ar = zip::ZipArchive::new(Cursor::new(raw)).unwrap();
        let entry = ar.by_name("word/document.xml").unwrap();
        // The size field is read from the header at open time — BEFORE any
        // decompression / CRC check. It is the attacker's declared value.
        assert_eq!(
            entry.size(),
            DECLARED as u64,
            "entry.size() must report the declared (attacker-controlled) size — \
             if this fails, the zip crate now returns the actual size and RB11's \
             reserve-inflation vector no longer exists"
        );
    }

    #[test]
    fn initial_reserve_caps_the_declared_size() {
        // The reserve helper clamps an entry's declared size to INITIAL_RESERVE_CAP
        // so a forged 512 MiB declaration reserves at most the cap, not 512 MiB.
        // A small declared size is reserved exactly (no waste for legitimate files).
        assert_eq!(
            initial_reserve(8),
            8,
            "a small declared size is reserved exactly"
        );
        assert_eq!(
            initial_reserve(INITIAL_RESERVE_CAP as u64),
            INITIAL_RESERVE_CAP,
            "a declared size equal to the cap reserves the cap"
        );
        assert_eq!(
            initial_reserve(512 * 1024 * 1024),
            INITIAL_RESERVE_CAP,
            "a forged 512 MiB declaration is clamped to the cap, not honored"
        );
    }

    #[test]
    fn read_zip_bytes_reads_full_data_despite_huge_declared_reserve() {
        // A legitimately-authored entry whose real body is small but sits in an
        // archive is read in FULL regardless of the reserve cap — the cap only
        // bounds the up-front `with_capacity`; `read_to_end` grows as needed.
        // (Uses a normal entry: correctness of the read path is what we assert;
        // the anti-over-reserve property is covered by initial_reserve_caps_*.)
        let mut ar = archive_with("word/document.xml", b"<document>hello</document>");
        assert_eq!(
            read_zip_bytes(&mut ar, "word/document.xml").unwrap(),
            b"<document>hello</document>",
            "read must yield the complete real body, reserve cap notwithstanding"
        );
    }

    #[test]
    fn read_zip_helpers_reject_oversized_under_scoped_cap() {
        // 8-byte entry; a scoped cap of 4 must reject (never truncate) — the
        // zip-bomb guard applies to the open-archive helpers too.
        let mut ar = archive_with("ppt/media/big.bin", b"12345678");
        {
            let _guard = scoped_max(Some(4));
            let be = read_zip_bytes(&mut ar, "ppt/media/big.bin").unwrap_err();
            assert!(be.contains("exceeds size limit"), "got: {be}");
            let se = read_zip_string(&mut ar, "ppt/media/big.bin").unwrap_err();
            assert!(se.contains("exceeds size limit"), "got: {se}");
        }
        // Cap restored on guard drop → the same entry now reads in full.
        assert_eq!(
            read_zip_bytes(&mut ar, "ppt/media/big.bin").unwrap(),
            b"12345678"
        );
    }
}
