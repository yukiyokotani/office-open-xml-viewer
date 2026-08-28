//! Owned, bounded access to one OOXML ZIP package.
//!
//! `zip::read::ZipFile` borrows its `ZipArchive`, which makes it unsuitable for
//! a reader that must survive multiple WASM/protocol pulls. This module uses the
//! ZIP crate to validate and index the package, then owns a decoder over the
//! selected compressed byte range. The original `Vec<u8>` allocation moves into
//! `Rc` unchanged; opening a reader does not copy the complete package.

use crate::resource::{
    HardResourceLimitKind, OoxmlFormat, ResourceGovernor, ResourceOperation, ResourceUsage,
};
use crate::{resource, zip as bounded_zip};
use crc32fast::Hasher;
use flate2::read::DeflateDecoder;
use std::cell::RefCell;
use std::collections::{HashMap, VecDeque};
use std::io::{Cursor, Read};
use std::rc::Rc;
use zip::CompressionMethod;

/// Raw entry payload returned by one bounded pull.
#[derive(Debug, PartialEq, Eq)]
pub(crate) struct EntryChunk {
    bytes: Vec<u8>,
    done: bool,
}

/// Opaque identity of one active bounded archive-entry reader.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub(crate) struct EntryReaderId(u64);

/// Backpressure ceiling for a single raw-entry pull. Resource limits remain
/// independent and may be lower.
pub(crate) const MAX_ENTRY_PULL_BYTES: usize = 1024 * 1024;
const MAX_ACTIVE_OPERATIONS: usize = 4096;
const HARD_MAX_ACTIVE_ENTRY_READERS: usize = 1024;
const MAX_FINALIZED_OPERATION_RECORDS: usize = 256;

#[derive(Clone)]
struct PackageBytes(Rc<Vec<u8>>);

impl PackageBytes {
    fn new(bytes: Vec<u8>) -> Self {
        Self(Rc::new(bytes))
    }

    fn as_slice(&self) -> &[u8] {
        self.0.as_slice()
    }
}

struct EntrySlice {
    source: PackageBytes,
    start: usize,
    position: usize,
    end: usize,
}

impl EntrySlice {
    fn new(source: PackageBytes, start: usize, end: usize) -> Self {
        Self {
            source,
            start,
            position: start,
            end,
        }
    }

    fn consumed(&self) -> u64 {
        self.position.saturating_sub(self.start) as u64
    }
}

impl Read for EntrySlice {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        let count = output.len().min(self.end.saturating_sub(self.position));
        if count == 0 {
            return Ok(0);
        }
        output[..count]
            .copy_from_slice(&self.source.as_slice()[self.position..self.position + count]);
        self.position += count;
        Ok(count)
    }
}

#[derive(Clone)]
struct EntryMetadata {
    index: usize,
    path: String,
    data_start: usize,
    data_end: usize,
    compressed_size: u64,
    declared_size: u64,
    crc32: u32,
    compression: CompressionMethod,
    encrypted: bool,
}

enum EntryDecoder {
    Stored(EntrySlice),
    Deflated(DeflateDecoder<EntrySlice>),
}

impl EntryDecoder {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        match self {
            Self::Stored(reader) => reader.read(output),
            Self::Deflated(reader) => reader.read(output),
        }
    }

    fn compressed_consumed(&self) -> u64 {
        match self {
            Self::Stored(reader) => reader.consumed(),
            Self::Deflated(reader) => reader.total_in(),
        }
    }
}

struct BoundedEntryReader {
    operation_id: ResourceOperation,
    metadata: EntryMetadata,
    decoder: EntryDecoder,
    crc: Hasher,
    actual_bytes: u64,
    lookahead: Option<u8>,
    finished: bool,
}

impl BoundedEntryReader {
    fn new(
        operation_id: ResourceOperation,
        metadata: EntryMetadata,
        source: PackageBytes,
    ) -> Result<Self, String> {
        if metadata.encrypted {
            return Err(format!(
                "encrypted ZIP entry is unsupported: {}",
                metadata.path
            ));
        }
        let slice = EntrySlice::new(source, metadata.data_start, metadata.data_end);
        let decoder = match metadata.compression {
            CompressionMethod::Stored => EntryDecoder::Stored(slice),
            CompressionMethod::Deflated => EntryDecoder::Deflated(DeflateDecoder::new(slice)),
            _ => {
                return Err(format!(
                    "unsupported ZIP compression method for entry: {}",
                    metadata.path
                ))
            }
        };
        Ok(Self {
            operation_id,
            metadata,
            decoder,
            crc: Hasher::new(),
            actual_bytes: 0,
            lookahead: None,
            finished: false,
        })
    }

    fn read_decoder(&mut self, output: &mut [u8]) -> Result<usize, String> {
        self.decoder
            .read(output)
            .map_err(|error| format!("ZIP entry decompression error: {error}"))
    }

    fn charge(&mut self, bytes: &[u8]) -> Result<(), String> {
        if bytes.is_empty() {
            return Ok(());
        }
        self.actual_bytes = self.actual_bytes.saturating_add(bytes.len() as u64);
        self.crc.update(bytes);
        resource::observe_inflated(
            self.metadata.index,
            &self.metadata.path,
            self.actual_bytes,
            bytes.len() as u64,
        )
    }

    fn finish(&mut self) -> Result<(), String> {
        if self.finished {
            return Ok(());
        }
        if self.decoder.compressed_consumed() != self.metadata.compressed_size {
            return Err(format!(
                "ZIP entry compressed-size mismatch: {}",
                self.metadata.path
            ));
        }
        if self.actual_bytes != self.metadata.declared_size {
            return Err(format!(
                "ZIP entry inflated-size mismatch: {}",
                self.metadata.path
            ));
        }
        let actual_crc = self.crc.clone().finalize();
        if actual_crc != self.metadata.crc32 {
            return Err(format!("ZIP entry CRC mismatch: {}", self.metadata.path));
        }
        self.finished = true;
        Ok(())
    }

    fn probe_after_full_chunk(&mut self) -> Result<bool, String> {
        let mut byte = [0u8; 1];
        let count = self.read_decoder(&mut byte)?;
        if count == 0 {
            self.finish()?;
            return Ok(true);
        }
        self.charge(&byte[..count])?;
        self.lookahead = Some(byte[0]);
        Ok(false)
    }

    fn pull(
        &mut self,
        credit: usize,
        resource_allowance: u64,
        probe_eof: bool,
    ) -> Result<EntryChunk, String> {
        if credit == 0 || credit > MAX_ENTRY_PULL_BYTES {
            return Err(format!(
                "entry pull credit must be between 1 and {MAX_ENTRY_PULL_BYTES} bytes"
            ));
        }
        if self.finished {
            return Ok(EntryChunk {
                bytes: Vec::new(),
                done: true,
            });
        }

        let mut output = Vec::with_capacity(credit);
        if let Some(byte) = self.lookahead.take() {
            output.push(byte);
        }
        let mut scratch = [0u8; 32 * 1024];

        while output.len() < credit {
            let resource_remaining = resource_allowance.saturating_sub(self.actual_bytes);
            if resource_remaining == 0 {
                // Distinguish exact EOF-at-limit from one proven byte beyond it.
                let mut byte = [0u8; 1];
                let count = self.read_decoder(&mut byte)?;
                if count == 0 {
                    self.finish()?;
                    return Ok(EntryChunk {
                        bytes: output,
                        done: true,
                    });
                }
                self.charge(&byte[..count])?;
                unreachable!("resource allowance + 1 must reject");
            }

            let wanted = (credit - output.len())
                .min(resource_remaining as usize)
                .min(scratch.len());
            let count = self.read_decoder(&mut scratch[..wanted])?;
            if count == 0 {
                self.finish()?;
                return Ok(EntryChunk {
                    bytes: output,
                    done: true,
                });
            }
            self.charge(&scratch[..count])?;
            output.extend_from_slice(&scratch[..count]);
        }

        let done = probe_eof && self.probe_after_full_chunk()?;
        Ok(EntryChunk {
            bytes: output,
            done,
        })
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OperationStatus {
    Active,
    Finished,
    Canceled,
}

struct OperationRecord {
    status: OperationStatus,
    final_usage: Option<ResourceUsage>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SessionState {
    Open,
    Closed,
}

/// One random-access OOXML package, its persistent resource governor, logical
/// operations, and active owned entry decoders.
pub(crate) struct PackageSession {
    source: Option<PackageBytes>,
    entries: Vec<EntryMetadata>,
    by_name: HashMap<String, usize>,
    governor: ResourceGovernor,
    operations: HashMap<ResourceOperation, OperationRecord>,
    finalized_operations: VecDeque<ResourceOperation>,
    readers: HashMap<EntryReaderId, BoundedEntryReader>,
    next_reader_id: u64,
    state: SessionState,
}

impl PackageSession {
    pub(crate) fn open(
        data: Vec<u8>,
        format: OoxmlFormat,
        max_archive_entry_bytes: Option<u64>,
        max_total_inflated_bytes: Option<u64>,
        max_archive_entries: Option<u64>,
    ) -> Result<Self, String> {
        let source = PackageBytes::new(data);
        let governor = ResourceGovernor::from_wasm(
            format,
            max_archive_entry_bytes,
            max_total_inflated_bytes,
            max_archive_entries,
        );
        let _scope = governor.scope("open");
        let mut archive = bounded_zip::open_validated_cursor(Cursor::new(source.as_slice()))?;

        let mut entries = Vec::with_capacity(archive.len());
        let mut by_name = HashMap::with_capacity(archive.len());
        for index in 0..archive.len() {
            let entry = archive
                .by_index_raw(index)
                .map_err(|error| format!("ZIP archive entry metadata error: {error}"))?;
            let data_start = usize::try_from(entry.data_start())
                .map_err(|_| "ZIP entry data offset exceeds this platform".to_string())?;
            let compressed_size = entry.compressed_size();
            let data_end = data_start
                .checked_add(
                    usize::try_from(compressed_size)
                        .map_err(|_| "ZIP entry size exceeds this platform".to_string())?,
                )
                .ok_or_else(|| "ZIP entry data range overflows this platform".to_string())?;
            if data_end > source.as_slice().len() {
                return Err("ZIP entry data range exceeds package bytes".to_string());
            }
            let path = entry.name().to_string();
            by_name.insert(path.clone(), index);
            entries.push(EntryMetadata {
                index,
                path,
                data_start,
                data_end,
                compressed_size,
                declared_size: entry.size(),
                crc32: entry.crc32(),
                compression: entry.compression(),
                encrypted: entry.encrypted(),
            });
        }
        drop(archive);
        drop(_scope);

        Ok(Self {
            source: Some(source),
            entries,
            by_name,
            governor,
            operations: HashMap::new(),
            finalized_operations: VecDeque::new(),
            readers: HashMap::new(),
            next_reader_id: 1,
            state: SessionState::Open,
        })
    }

    fn ensure_healthy(&self) -> Result<(), String> {
        if self.state == SessionState::Closed {
            return Err("package session is closed".to_string());
        }
        if let Some(error) = self.governor.first_error() {
            return Err(error);
        }
        Ok(())
    }

    fn assert_operation_active(&self, id: ResourceOperation) -> Result<(), String> {
        let operation = self
            .operations
            .get(&id)
            .ok_or_else(|| format!("unknown package operation: {id:?}"))?;
        if operation.status != OperationStatus::Active {
            return Err(format!("package operation is not active: {id:?}"));
        }
        Ok(())
    }

    fn converge_poison(&mut self) {
        if self.governor.first_error().is_none() {
            return;
        }
        self.readers.clear();
        let mut finalized = Vec::new();
        for (id, operation) in &mut self.operations {
            if operation.status == OperationStatus::Active {
                operation.final_usage = self.governor.usage_for_operation(*id);
                operation.status = OperationStatus::Canceled;
                finalized.push(*id);
            }
        }
        self.governor.clear_operations();
        for id in finalized {
            self.remember_finalized(id);
        }
    }

    fn remember_finalized(&mut self, operation_id: ResourceOperation) {
        self.finalized_operations.push_back(operation_id);
        while self.finalized_operations.len() > MAX_FINALIZED_OPERATION_RECORDS {
            if let Some(expired) = self.finalized_operations.pop_front() {
                self.operations.remove(&expired);
            }
        }
    }

    pub(crate) fn begin_operation(
        &mut self,
        name: impl Into<String>,
    ) -> Result<ResourceOperation, String> {
        self.ensure_healthy()?;
        if self
            .operations
            .values()
            .filter(|operation| operation.status == OperationStatus::Active)
            .count()
            >= MAX_ACTIVE_OPERATIONS
        {
            return Err(format!(
                "package session cannot exceed {MAX_ACTIVE_OPERATIONS} active operations"
            ));
        }
        let id = self.governor.begin_operation(name);
        self.operations.insert(
            id,
            OperationRecord {
                status: OperationStatus::Active,
                final_usage: None,
            },
        );
        Ok(id)
    }

    pub(crate) fn open_entry(
        &mut self,
        operation_id: ResourceOperation,
        path: &str,
    ) -> Result<EntryReaderId, String> {
        self.ensure_healthy()?;
        self.assert_operation_active(operation_id)?;
        let index = self
            .by_name
            .get(path)
            .copied()
            .ok_or_else(|| format!("entry not found: {path}"))?;
        self.open_entry_by_index(operation_id, index)
    }

    pub(crate) fn open_entry_by_index(
        &mut self,
        operation_id: ResourceOperation,
        index: usize,
    ) -> Result<EntryReaderId, String> {
        self.ensure_healthy()?;
        self.assert_operation_active(operation_id)?;
        if self.readers.len() >= HARD_MAX_ACTIVE_ENTRY_READERS {
            return Err(format!(
                "package session cannot exceed {HARD_MAX_ACTIVE_ENTRY_READERS} active entry readers"
            ));
        }
        let metadata = self
            .entries
            .get(index)
            .cloned()
            .ok_or_else(|| format!("archive entry index out of range: {index}"))?;
        let source = self
            .source
            .as_ref()
            .cloned()
            .ok_or_else(|| "package session is closed".to_string())?;
        let scope = self.governor.scope_operation(operation_id)?;
        let result =
            resource::read_allowance(metadata.index, &metadata.path, metadata.declared_size)
                .and_then(|_| BoundedEntryReader::new(operation_id, metadata, source));
        drop(scope);
        let reader = match result {
            Ok(reader) => reader,
            Err(error) => {
                self.converge_poison();
                return Err(error);
            }
        };
        let id = EntryReaderId(self.next_reader_id);
        let next_reader_id = self
            .next_reader_id
            .checked_add(1)
            .ok_or_else(|| "package entry reader ID space is exhausted".to_string())?;
        if self.readers.contains_key(&id) {
            return Err("package entry reader ID space is exhausted".to_string());
        }
        self.next_reader_id = next_reader_id;
        self.readers.insert(id, reader);
        Ok(id)
    }

    pub(crate) fn pull_entry(
        &mut self,
        operation_id: ResourceOperation,
        reader_id: EntryReaderId,
        max_bytes: usize,
    ) -> Result<EntryChunk, String> {
        self.pull_entry_with_eof_probe(operation_id, reader_id, max_bytes, true)
    }

    fn pull_entry_prefix(
        &mut self,
        operation_id: ResourceOperation,
        reader_id: EntryReaderId,
        max_bytes: usize,
    ) -> Result<EntryChunk, String> {
        self.pull_entry_with_eof_probe(operation_id, reader_id, max_bytes, false)
    }

    fn pull_entry_with_eof_probe(
        &mut self,
        operation_id: ResourceOperation,
        reader_id: EntryReaderId,
        max_bytes: usize,
        probe_eof: bool,
    ) -> Result<EntryChunk, String> {
        self.ensure_healthy()?;
        if max_bytes == 0 || max_bytes > MAX_ENTRY_PULL_BYTES {
            return Err(format!(
                "entry pull credit must be between 1 and {MAX_ENTRY_PULL_BYTES} bytes"
            ));
        }
        self.assert_operation_active(operation_id)?;
        let mut reader = self
            .readers
            .remove(&reader_id)
            .ok_or_else(|| format!("unknown package entry reader: {}", reader_id.0))?;
        if reader.operation_id != operation_id {
            self.readers.insert(reader_id, reader);
            return Err("entry reader belongs to another operation".to_string());
        }

        let scope = self.governor.scope_operation(operation_id)?;
        // Re-evaluate the shared session budget for every pull. Readers may be
        // interleaved, so an allowance captured when an entry was opened would
        // become stale after another entry consumes distinct-inflated bytes.
        // `&mut self` and the synchronous decoder loop make this allowance
        // check plus all charges one non-interleavable package transaction.
        let result = resource::read_allowance(
            reader.metadata.index,
            &reader.metadata.path,
            reader.metadata.declared_size,
        )
        .and_then(|allowance| reader.pull(max_bytes, allowance, probe_eof));
        drop(scope);

        match result {
            Ok(chunk) => {
                if !chunk.done {
                    self.readers.insert(reader_id, reader);
                }
                Ok(chunk)
            }
            Err(error) => {
                // A resource violation poisons the package as a whole. Drop
                // every decoder immediately so sibling operations do not
                // retain compressed input or decompression state.
                self.converge_poison();
                Err(error)
            }
        }
    }

    pub(crate) fn release_entry(&mut self, reader_id: EntryReaderId) {
        self.readers.remove(&reader_id);
    }

    /// Number of entries exposed by the validated, uniquely named ZIP index.
    pub(crate) fn entry_count(&self) -> usize {
        self.entries.len()
    }

    /// Whether a validated package contains this exact entry path.
    pub(crate) fn contains_entry(&self, path: &str) -> bool {
        self.by_name.contains_key(path)
    }

    /// Clone entry paths in deterministic validated-index order without
    /// exposing decoder metadata or borrowing the session internals.
    pub(crate) fn entry_paths(&self) -> Vec<String> {
        self.entries
            .iter()
            .map(|entry| entry.path.clone())
            .collect()
    }

    pub(crate) fn finish_operation(
        &mut self,
        operation_id: ResourceOperation,
    ) -> Result<(), String> {
        self.ensure_healthy()?;
        let Some(operation) = self.operations.get_mut(&operation_id) else {
            return Err(format!("unknown package operation: {operation_id:?}"));
        };
        if operation.status == OperationStatus::Active {
            operation.final_usage = self.governor.usage_for_operation(operation_id);
            self.governor.end_operation(operation_id);
            operation.status = OperationStatus::Finished;
            self.readers
                .retain(|_, reader| reader.operation_id != operation_id);
            self.remember_finalized(operation_id);
        }
        Ok(())
    }

    pub(crate) fn cancel_operation(
        &mut self,
        operation_id: ResourceOperation,
    ) -> Result<(), String> {
        if self.state == SessionState::Closed {
            return Ok(());
        }
        let Some(operation) = self.operations.get_mut(&operation_id) else {
            return Err(format!("unknown package operation: {operation_id:?}"));
        };
        if operation.status == OperationStatus::Active {
            operation.final_usage = self.governor.usage_for_operation(operation_id);
            self.governor.end_operation(operation_id);
            operation.status = OperationStatus::Canceled;
            self.readers
                .retain(|_, reader| reader.operation_id != operation_id);
            self.remember_finalized(operation_id);
        }
        Ok(())
    }

    pub(crate) fn close(&mut self) {
        if self.state == SessionState::Closed {
            return;
        }
        self.readers.clear();
        self.operations.clear();
        self.finalized_operations.clear();
        self.governor.clear_operations();
        self.entries.clear();
        self.by_name.clear();
        self.source = None;
        self.state = SessionState::Closed;
    }

    pub(crate) fn usage(&self) -> ResourceUsage {
        self.governor.usage()
    }

    /// Return a checkpoint with the session-wide counters and the selected
    /// logical operation's cumulative work, even when another operation pulled
    /// most recently.
    pub(crate) fn usage_for_operation(
        &self,
        operation_id: ResourceOperation,
    ) -> Option<ResourceUsage> {
        let operation = self.operations.get(&operation_id)?;
        operation
            .final_usage
            .or_else(|| self.governor.usage_for_operation(operation_id))
    }

    fn observe_hard_limit(
        &mut self,
        operation_id: ResourceOperation,
        kind: HardResourceLimitKind,
        part: Option<&str>,
        limit: u64,
        observed: u64,
    ) -> Result<(), String> {
        self.ensure_healthy()?;
        self.assert_operation_active(operation_id)?;
        let scope = self.governor.scope_operation(operation_id)?;
        let result = resource::observe_hard_limit(kind, part, limit, observed);
        drop(scope);
        if result.is_err() {
            self.converge_poison();
        }
        result
    }

    #[cfg(test)]
    fn operation_inflated_bytes(&self, operation_id: ResourceOperation) -> Option<u64> {
        self.usage_for_operation(operation_id)
            .map(|usage| usage.operation_inflated_bytes)
    }
}

impl Drop for PackageSession {
    fn drop(&mut self) {
        self.close();
    }
}

/// Cloneable, single-threaded ownership boundary for one OOXML package.
///
/// Every clone points at the same source bytes, decoder registry, and resource
/// governor. Lifecycle mutation stays behind RAII [`PackageOperation`] and
/// [`PackageEntryStream`] values rather than being exposed as raw identifiers
/// to parser code.
#[derive(Clone)]
pub struct PackageSessionHandle {
    inner: Rc<RefCell<PackageSession>>,
}

impl PackageSessionHandle {
    pub fn open(
        data: Vec<u8>,
        format: OoxmlFormat,
        max_archive_entry_bytes: Option<u64>,
        max_total_inflated_bytes: Option<u64>,
        max_archive_entries: Option<u64>,
    ) -> Result<Self, String> {
        PackageSession::open(
            data,
            format,
            max_archive_entry_bytes,
            max_total_inflated_bytes,
            max_archive_entries,
        )
        .map(Self::new)
    }

    pub(crate) fn new(session: PackageSession) -> Self {
        Self {
            inner: Rc::new(RefCell::new(session)),
        }
    }

    pub fn begin_operation(&self, name: impl Into<String>) -> Result<PackageOperation, String> {
        let id = self.inner.borrow_mut().begin_operation(name)?;
        Ok(PackageOperation {
            handle: self.clone(),
            id,
            state: OwnedOperationState::Active,
        })
    }

    /// Run one self-contained package read in its own accounting operation.
    ///
    /// Format cursors may retain a different operation across an acknowledged
    /// unit boundary. This helper deliberately does not borrow that cursor's
    /// operation: the package session owns both lifecycles and applies the same
    /// shared governor to them. Success commits only this operation; failure
    /// cancels only this operation while preserving a package-wide resource
    /// failure as the primary error.
    pub fn run_operation<T>(
        &self,
        name: impl Into<String>,
        run: impl FnOnce(&PackageOperation) -> Result<T, String>,
    ) -> Result<T, String> {
        let mut operation = self.begin_operation(name)?;
        let result = run(&operation);
        if let Err(resource_error) = self.assert_healthy() {
            let _ = operation.cancel();
            return Err(resource_error);
        }
        match result {
            Ok(value) => match operation.finish() {
                Ok(()) => Ok(value),
                Err(error) => {
                    let _ = operation.cancel();
                    Err(error)
                }
            },
            Err(error) => {
                let _ = operation.cancel();
                Err(error)
            }
        }
    }

    /// Verify that the package is still open and has not encountered a
    /// package-wide resource failure.
    pub fn assert_healthy(&self) -> Result<(), String> {
        self.inner.borrow().ensure_healthy()
    }

    pub fn usage(&self) -> ResourceUsage {
        self.inner.borrow().usage()
    }

    pub fn entry_count(&self) -> usize {
        self.inner.borrow().entry_count()
    }

    pub fn contains_entry(&self, path: &str) -> bool {
        self.inner.borrow().contains_entry(path)
    }

    pub fn entry_paths(&self) -> Vec<String> {
        self.inner.borrow().entry_paths()
    }

    /// Close the shared package. All handle clones, operations, and entry
    /// streams observe the same closed state.
    pub fn close(&self) {
        self.inner.borrow_mut().close();
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OwnedOperationState {
    Active,
    Finished,
    Canceled,
}

/// RAII owner of one logical resource-accounting operation.
///
/// Dropping an active operation cancels it and releases every reader opened by
/// that operation. Explicit finish and cancel calls are idempotent.
pub struct PackageOperation {
    handle: PackageSessionHandle,
    id: ResourceOperation,
    state: OwnedOperationState,
}

impl PackageOperation {
    /// Derive format-model limit reporting from this logical operation. This is
    /// used by materializing compatibility paths that do not have an entry
    /// stream in hand when their retained-model or serialization ceiling is
    /// crossed.
    pub fn limit_reporter(&self) -> Result<PackageLimitReporter, String> {
        let session = self.handle.inner.borrow();
        session.ensure_healthy()?;
        session.assert_operation_active(self.id)?;
        Ok(PackageLimitReporter {
            handle: self.handle.clone(),
            operation_id: self.id,
        })
    }

    fn assert_active(&self) -> Result<(), String> {
        if self.state != OwnedOperationState::Active {
            return Err("package operation is not active".to_string());
        }
        let session = self.handle.inner.borrow();
        session.ensure_healthy()?;
        session.assert_operation_active(self.id)
    }

    pub fn open_entry(&self, path: &str) -> Result<PackageEntryStream, String> {
        self.assert_active()?;
        let reader_id = self.handle.inner.borrow_mut().open_entry(self.id, path)?;
        Ok(PackageEntryStream {
            handle: self.handle.clone(),
            operation_id: self.id,
            reader_id: Some(reader_id),
            part: path.to_string(),
            done: false,
            eof_probe: true,
        })
    }

    pub fn open_entry_by_index(&self, index: usize) -> Result<PackageEntryStream, String> {
        self.assert_active()?;
        let part = self
            .handle
            .inner
            .borrow()
            .entries
            .get(index)
            .map(|entry| entry.path.clone())
            .ok_or_else(|| format!("ZIP archive entry index is out of range: {index}"))?;
        let reader_id = self
            .handle
            .inner
            .borrow_mut()
            .open_entry_by_index(self.id, index)?;
        Ok(PackageEntryStream {
            handle: self.handle.clone(),
            operation_id: self.id,
            reader_id: Some(reader_id),
            part,
            done: false,
            eof_probe: true,
        })
    }

    fn open_entry_prefix(&self, path: &str) -> Result<PackageEntryStream, String> {
        self.assert_active()?;
        let reader_id = self.handle.inner.borrow_mut().open_entry(self.id, path)?;
        Ok(PackageEntryStream {
            handle: self.handle.clone(),
            operation_id: self.id,
            reader_id: Some(reader_id),
            part: path.to_string(),
            done: false,
            eof_probe: false,
        })
    }

    /// Materialize one entry through the same bounded decoder and accounting
    /// path used by streaming consumers.
    pub fn read_bytes(&self, path: &str) -> Result<Vec<u8>, String> {
        let mut stream = self.open_entry(path)?;
        let mut bytes = Vec::new();
        stream
            .read_to_end(&mut bytes)
            .map_err(|error| error.to_string())?;
        Ok(bytes)
    }

    /// Materialize one entry without ever retaining more than `max_bytes + 1`.
    /// The validated central-directory size provides an allocation-free early
    /// rejection; the limit+1 decoder read also defends against a corrupt stream
    /// whose actual output exceeds that declaration. This is an optional-part
    /// admission check, not a package-wide resource poison.
    pub fn read_bytes_bounded(&self, path: &str, max_bytes: usize) -> Result<Vec<u8>, String> {
        self.assert_active()?;
        let declared_size = {
            let session = self.handle.inner.borrow();
            let index = session
                .by_name
                .get(path)
                .copied()
                .ok_or_else(|| format!("ZIP entry not found: {path}"))?;
            session.entries[index].declared_size
        };
        if declared_size > max_bytes as u64 {
            return Err(format!(
                "ZIP entry exceeds optional-part byte limit ({path}): {declared_size} > {max_bytes}"
            ));
        }

        let retained_cap = max_bytes
            .checked_add(1)
            .ok_or_else(|| "optional-part byte limit overflows this platform".to_string())?;
        let mut stream = self.open_entry(path)?;
        let mut bytes = Vec::with_capacity(declared_size.min(max_bytes as u64) as usize);
        let mut scratch = [0u8; 8 * 1024];
        while bytes.len() < retained_cap {
            let wanted = (retained_cap - bytes.len()).min(scratch.len());
            let count = stream
                .read(&mut scratch[..wanted])
                .map_err(|error| error.to_string())?;
            if count == 0 {
                break;
            }
            bytes.extend_from_slice(&scratch[..count]);
        }
        if bytes.len() > max_bytes {
            return Err(format!(
                "ZIP entry exceeds optional-part byte limit ({path}): more than {max_bytes} bytes"
            ));
        }
        Ok(bytes)
    }

    pub fn read_string(&self, path: &str) -> Result<String, String> {
        String::from_utf8(self.read_bytes(path)?)
            .map_err(|error| format!("ZIP entry is not valid UTF-8 ({path}): {error}"))
    }

    /// Read at most `max_bytes` from the start of an entry. Dropping the local
    /// stream releases its decoder. Prefix reads deliberately skip the normal
    /// one-byte EOF probe, so accounting includes only bytes returned here.
    pub fn read_head(&self, path: &str, max_bytes: usize) -> Result<Vec<u8>, String> {
        // Open even for an empty prefix so path existence, operation state,
        // session health, and declared entry limits remain consistently
        // validated without inflating any entry bytes.
        let mut stream = self.open_entry_prefix(path)?;
        if max_bytes == 0 {
            return Ok(Vec::new());
        }
        let mut bytes = Vec::with_capacity(max_bytes.min(32 * 1024));
        let mut scratch = [0u8; 8 * 1024];
        while bytes.len() < max_bytes {
            let wanted = (max_bytes - bytes.len()).min(scratch.len());
            let count = stream
                .read(&mut scratch[..wanted])
                .map_err(|error| error.to_string())?;
            if count == 0 {
                break;
            }
            bytes.extend_from_slice(&scratch[..count]);
        }
        Ok(bytes)
    }

    pub fn usage(&self) -> Option<ResourceUsage> {
        self.handle.inner.borrow().usage_for_operation(self.id)
    }

    pub fn finish(&mut self) -> Result<(), String> {
        if self.state != OwnedOperationState::Active {
            return Ok(());
        }
        self.handle.inner.borrow_mut().finish_operation(self.id)?;
        self.state = OwnedOperationState::Finished;
        Ok(())
    }

    pub fn cancel(&mut self) -> Result<(), String> {
        if self.state != OwnedOperationState::Active {
            return Ok(());
        }
        self.handle.inner.borrow_mut().cancel_operation(self.id)?;
        self.state = OwnedOperationState::Canceled;
        Ok(())
    }
}

impl Drop for PackageOperation {
    fn drop(&mut self) {
        let _ = self.cancel();
    }
}

/// Owns the single retained package operation used by a format parser cursor.
///
/// DOCX, XLSX, and PPTX expose different parser models, but their package-read
/// lifecycle is identical: one explicitly started operation remains active
/// across the format-specific work and is then committed or canceled as one
/// unit. Keeping that ownership protocol here prevents the three adapters from
/// drifting in health checks, error precedence, or cleanup behavior.
pub struct RetainedPackageOperation {
    format: &'static str,
    operation: Option<PackageOperation>,
}

impl RetainedPackageOperation {
    pub fn new(format: &'static str) -> Self {
        Self {
            format,
            operation: None,
        }
    }

    pub fn begin(&mut self, session: &PackageSessionHandle, name: &str) -> Result<(), String> {
        if self.operation.is_some() {
            return Err(format!(
                "{} package operation is already active",
                self.format
            ));
        }
        self.operation = Some(session.begin_operation(name)?);
        Ok(())
    }

    /// Return the retained operation, optionally creating a test-only
    /// compatibility scope selected by the format adapter.
    pub fn operation(
        &mut self,
        session: &PackageSessionHandle,
        compatibility_name: Option<&str>,
    ) -> Result<&PackageOperation, String> {
        if self.operation.is_none() {
            let Some(name) = compatibility_name else {
                return Err(format!(
                    "{} package read requires an active operation",
                    self.format
                ));
            };
            self.operation = Some(session.begin_operation(name)?);
        }
        Ok(self
            .operation
            .as_ref()
            .expect("operation initialized above"))
    }

    pub fn active(&self) -> Result<&PackageOperation, String> {
        self.operation
            .as_ref()
            .ok_or_else(|| format!("{} package operation is not active", self.format))
    }

    pub fn is_active(&self) -> bool {
        self.operation.is_some()
    }

    pub fn usage(&self) -> Option<ResourceUsage> {
        self.operation.as_ref().and_then(PackageOperation::usage)
    }

    pub fn finish(&mut self) -> Result<(), String> {
        let Some(mut operation) = self.operation.take() else {
            return Ok(());
        };
        operation.finish()
    }

    pub fn cancel(&mut self) {
        if let Some(mut operation) = self.operation.take() {
            let _ = operation.cancel();
        }
    }

    /// Settle a format operation after its format-specific work finishes.
    /// Package-wide resource failures take precedence over parse/model errors.
    pub fn settle<T>(
        &mut self,
        session: &PackageSessionHandle,
        result: Result<T, String>,
    ) -> Result<T, String> {
        if let Err(resource_error) = session.assert_healthy() {
            self.cancel();
            return Err(resource_error);
        }
        match result {
            Ok(value) => match self.finish() {
                Ok(()) => Ok(value),
                Err(error) => {
                    self.cancel();
                    Err(error)
                }
            },
            Err(error) => {
                self.cancel();
                Err(error)
            }
        }
    }
}

/// Operation-bound reporting capability for parser/model safety ceilings.
///
/// This capability does not own operation lifecycle. Dropping it is inert, and
/// using it after the operation finishes or cancels is rejected.
#[derive(Clone)]
pub struct PackageLimitReporter {
    handle: PackageSessionHandle,
    operation_id: ResourceOperation,
}

impl PackageLimitReporter {
    pub fn observe_hard_limit(
        &self,
        kind: HardResourceLimitKind,
        part: Option<&str>,
        limit: u64,
        observed: u64,
    ) -> Result<(), String> {
        self.handle.inner.borrow_mut().observe_hard_limit(
            self.operation_id,
            kind,
            part,
            limit,
            observed,
        )
    }
}

/// Owned, resumable archive-entry reader suitable for `quick_xml::Reader`.
///
/// A `Read` call never asks the underlying package session for more than
/// the internal 1 MiB pull ceiling, even when the caller supplies a larger
/// buffer.
pub struct PackageEntryStream {
    handle: PackageSessionHandle,
    operation_id: ResourceOperation,
    reader_id: Option<EntryReaderId>,
    part: String,
    done: bool,
    eof_probe: bool,
}

impl PackageEntryStream {
    pub fn part_name(&self) -> &str {
        &self.part
    }

    /// Derive parser-limit reporting from the entry stream itself so inflated
    /// bytes and parser/model ceilings cannot be attributed to different
    /// package operations.
    pub fn limit_reporter(&self) -> Result<PackageLimitReporter, String> {
        let session = self.handle.inner.borrow();
        session.ensure_healthy()?;
        session.assert_operation_active(self.operation_id)?;
        Ok(PackageLimitReporter {
            handle: self.handle.clone(),
            operation_id: self.operation_id,
        })
    }
}

impl Read for PackageEntryStream {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        if output.is_empty() || self.done {
            return Ok(0);
        }
        let Some(reader_id) = self.reader_id else {
            return Ok(0);
        };
        let credit = output.len().min(MAX_ENTRY_PULL_BYTES);
        let chunk = match if self.eof_probe {
            self.handle
                .inner
                .borrow_mut()
                .pull_entry(self.operation_id, reader_id, credit)
        } else {
            self.handle
                .inner
                .borrow_mut()
                .pull_entry_prefix(self.operation_id, reader_id, credit)
        } {
            Ok(chunk) => chunk,
            Err(error) => {
                // `PackageSession::pull_entry` consumes a failed decoder.
                self.reader_id = None;
                return Err(std::io::Error::other(error));
            }
        };
        let count = chunk.bytes.len();
        output[..count].copy_from_slice(&chunk.bytes);
        if chunk.done {
            // A terminal chunk can still contain bytes. Return those bytes now
            // and report EOF on the next call, as required by `Read`.
            self.done = true;
            self.reader_id = None;
        }
        Ok(count)
    }
}

impl Drop for PackageEntryStream {
    fn drop(&mut self) {
        if let Some(reader_id) = self.reader_id.take() {
            self.handle.inner.borrow_mut().release_entry(reader_id);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::write::SimpleFileOptions;

    fn package(entries: &[(&str, &[u8], CompressionMethod)]) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            for (name, body, compression) in entries {
                writer
                    .start_file(
                        *name,
                        SimpleFileOptions::default().compression_method(*compression),
                    )
                    .unwrap();
                writer.write_all(body).unwrap();
            }
            writer.finish().unwrap();
        }
        bytes
    }

    fn drain(
        session: &mut PackageSession,
        operation: ResourceOperation,
        reader: EntryReaderId,
        credit: usize,
    ) -> Result<Vec<u8>, String> {
        let mut all = Vec::new();
        loop {
            let chunk = session.pull_entry(operation, reader, credit)?;
            assert!(chunk.bytes.len() <= credit);
            all.extend_from_slice(&chunk.bytes);
            if chunk.done {
                return Ok(all);
            }
        }
    }

    #[test]
    fn retained_operation_requires_explicit_scope_unless_adapter_allows_compatibility() {
        let bytes = package(&[("word/document.xml", b"<w/>", CompressionMethod::Stored)]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Docx, Some(64), Some(64), None).unwrap();
        let mut retained = RetainedPackageOperation::new("docx");

        assert_eq!(
            retained.operation(&handle, None).err().unwrap(),
            "docx package read requires an active operation"
        );
        assert_eq!(
            retained
                .operation(&handle, Some("docx-parser-compat"))
                .unwrap()
                .read_bytes("word/document.xml")
                .unwrap(),
            b"<w/>"
        );
        assert!(retained.is_active());
        retained.cancel();
        assert!(!retained.is_active());
    }

    #[test]
    fn retained_operation_settlement_prioritizes_package_resource_failure() {
        let bytes = package(&[(
            "xl/worksheets/sheet1.xml",
            b"<x/>",
            CompressionMethod::Stored,
        )]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Xlsx, Some(64), Some(64), None).unwrap();
        let mut retained = RetainedPackageOperation::new("xlsx");
        retained.begin(&handle, "parse-sheet").unwrap();
        let reporter = retained.active().unwrap().limit_reporter().unwrap();
        let resource_error = reporter
            .observe_hard_limit(
                HardResourceLimitKind::XmlEventBytes,
                Some("xl/worksheets/sheet1.xml"),
                8,
                9,
            )
            .unwrap_err();

        assert_eq!(
            retained
                .settle::<()>(&handle, Err("secondary parse failure".to_string()))
                .unwrap_err(),
            resource_error
        );
        assert!(!retained.is_active());
    }

    #[test]
    fn stored_and_deflated_entries_cross_bounded_pulls() {
        let body = b"abcdefghijklmnopqrstuvwxyz";
        for compression in [CompressionMethod::Stored, CompressionMethod::Deflated] {
            let bytes = package(&[("word/document.xml", body, compression)]);
            let mut session =
                PackageSession::open(bytes, OoxmlFormat::Docx, Some(1024), Some(1024), None)
                    .unwrap();
            let operation = session.begin_operation("parse").unwrap();
            let reader = session.open_entry(operation, "word/document.xml").unwrap();
            assert_eq!(drain(&mut session, operation, reader, 5).unwrap(), body);
            assert_eq!(session.operation_inflated_bytes(operation), Some(26));
        }
    }

    #[test]
    fn bounded_materialization_rejects_before_retaining_an_oversized_optional_part() {
        for compression in [CompressionMethod::Stored, CompressionMethod::Deflated] {
            let bytes = package(&[("ppt/fonts/font1.fntdata", b"12345", compression)]);
            let handle =
                PackageSessionHandle::open(bytes, OoxmlFormat::Pptx, Some(64), Some(64), None)
                    .unwrap();
            let operation = handle.begin_operation("extract-font").unwrap();
            assert_eq!(
                operation
                    .read_bytes_bounded("ppt/fonts/font1.fntdata", 5)
                    .unwrap(),
                b"12345"
            );
            assert!(operation
                .read_bytes_bounded("ppt/fonts/font1.fntdata", 4)
                .unwrap_err()
                .contains("optional-part byte limit"));
        }
    }

    #[test]
    fn validated_archive_offset_is_used_for_prefixed_packages() {
        let package = package(&[(
            "xl/workbook.xml",
            b"prefixed package",
            CompressionMethod::Deflated,
        )]);
        let mut prefixed = b"MZ\x90\0self-extracting-prefix".to_vec();
        prefixed.extend_from_slice(&package);
        let handle =
            PackageSessionHandle::open(prefixed, OoxmlFormat::Xlsx, Some(64), Some(64), None)
                .unwrap();
        let operation = handle.begin_operation("parse").unwrap();
        assert_eq!(
            operation.read_bytes("xl/workbook.xml").unwrap(),
            b"prefixed package"
        );
    }

    #[test]
    fn full_credit_at_eof_is_reported_done_without_an_empty_pull() {
        let bytes = package(&[(
            "xl/worksheets/sheet1.xml",
            b"12345678",
            CompressionMethod::Deflated,
        )]);
        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Xlsx, Some(8), Some(8), None).unwrap();
        let operation = session.begin_operation("sheet").unwrap();
        let reader = session
            .open_entry(operation, "xl/worksheets/sheet1.xml")
            .unwrap();
        let chunk = session.pull_entry(operation, reader, 8).unwrap();
        assert_eq!(chunk.bytes, b"12345678");
        assert!(chunk.done);
    }

    #[test]
    fn pull_credit_is_positive_and_internally_bounded() {
        let bytes = package(&[("word/a.xml", b"1", CompressionMethod::Stored)]);
        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Docx, Some(8), Some(8), None).unwrap();
        let operation = session.begin_operation("parse").unwrap();
        let reader = session.open_entry(operation, "word/a.xml").unwrap();
        assert!(session.pull_entry(operation, reader, 0).is_err());
        assert!(session
            .pull_entry(operation, reader, MAX_ENTRY_PULL_BYTES + 1)
            .is_err());
        assert_eq!(
            session.pull_entry(operation, reader, 1).unwrap(),
            EntryChunk {
                bytes: b"1".to_vec(),
                done: true,
            }
        );
    }

    #[test]
    fn lookahead_is_delivered_once_and_operation_usage_survives_pulls() {
        let bytes = package(&[(
            "ppt/slides/slide1.xml",
            b"1234567",
            CompressionMethod::Stored,
        )]);
        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Pptx, Some(64), Some(64), None).unwrap();
        let operation = session.begin_operation("slide").unwrap();
        let reader = session
            .open_entry(operation, "ppt/slides/slide1.xml")
            .unwrap();
        let first = session.pull_entry(operation, reader, 3).unwrap();
        assert_eq!(first.bytes, b"123");
        assert!(!first.done);
        // One-byte lookahead was decompressed and charged, but is not duplicated.
        assert_eq!(session.operation_inflated_bytes(operation), Some(4));
        let rest = drain(&mut session, operation, reader, 3).unwrap();
        assert_eq!(rest, b"4567");
        assert_eq!(session.operation_inflated_bytes(operation), Some(7));
    }

    #[test]
    fn exact_limit_succeeds_and_limit_plus_one_poisons_every_later_read() {
        let exact = package(&[("word/a.xml", b"1234", CompressionMethod::Stored)]);
        let mut session =
            PackageSession::open(exact, OoxmlFormat::Docx, Some(4), Some(4), None).unwrap();
        let operation = session.begin_operation("parse").unwrap();
        let reader = session.open_entry(operation, "word/a.xml").unwrap();
        assert_eq!(drain(&mut session, operation, reader, 2).unwrap(), b"1234");

        let mut over = package(&[
            ("word/a.xml", b"12345", CompressionMethod::Stored),
            ("word/b.xml", b"x", CompressionMethod::Stored),
        ]);
        // Forge the first entry's local and central uncompressed declarations
        // below the policy. Its stored payload still emits five actual bytes.
        over[22..26].copy_from_slice(&1u32.to_le_bytes());
        let central = over
            .windows(4)
            .position(|window| window == 0x0201_4b50u32.to_le_bytes())
            .unwrap();
        over[central + 24..central + 28].copy_from_slice(&1u32.to_le_bytes());
        let mut session =
            PackageSession::open(over, OoxmlFormat::Docx, Some(4), Some(64), None).unwrap();
        let operation = session.begin_operation("parse").unwrap();
        let reader = session.open_entry(operation, "word/a.xml").unwrap();
        let sibling = session.open_entry(operation, "word/b.xml").unwrap();
        let first = drain(&mut session, operation, reader, 2).unwrap_err();
        assert!(first.contains("actual-inflated-bytes"));
        assert!(session.readers.is_empty());
        assert!(!session.readers.contains_key(&sibling));
        assert!(session.usage_for_operation(operation).is_some());
        assert_eq!(
            session.open_entry(operation, "word/b.xml").unwrap_err(),
            first
        );
        assert_eq!(
            session.open_entry(operation, "missing.xml").unwrap_err(),
            first
        );
    }

    #[test]
    fn reread_counts_operation_work_but_not_distinct_session_bytes() {
        let bytes = package(&[("word/a.xml", b"1234", CompressionMethod::Deflated)]);
        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Docx, Some(16), Some(4), None).unwrap();
        let operation = session.begin_operation("parse-and-markdown").unwrap();
        for _ in 0..2 {
            let reader = session.open_entry(operation, "word/a.xml").unwrap();
            assert_eq!(drain(&mut session, operation, reader, 2).unwrap(), b"1234");
        }
        assert_eq!(session.operation_inflated_bytes(operation), Some(8));
        assert_eq!(session.usage().distinct_inflated_bytes, 4);
    }

    #[test]
    fn interleaved_operations_keep_independent_work_checkpoints() {
        let bytes = package(&[
            ("word/a.xml", b"1234", CompressionMethod::Stored),
            ("word/b.xml", b"567", CompressionMethod::Stored),
        ]);
        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Docx, Some(16), Some(16), None).unwrap();
        let first_operation = session.begin_operation("first").unwrap();
        let second_operation = session.begin_operation("second").unwrap();
        let first_reader = session.open_entry(first_operation, "word/a.xml").unwrap();
        let second_reader = session.open_entry(second_operation, "word/b.xml").unwrap();

        assert_eq!(
            drain(&mut session, first_operation, first_reader, 2).unwrap(),
            b"1234"
        );
        assert_eq!(
            drain(&mut session, second_operation, second_reader, 2).unwrap(),
            b"567"
        );
        assert_eq!(session.operation_inflated_bytes(first_operation), Some(4));
        assert_eq!(session.operation_inflated_bytes(second_operation), Some(3));
        assert_eq!(session.usage().distinct_inflated_bytes, 7);
    }

    #[test]
    fn central_directory_indices_keep_distinct_accounting() {
        let bytes = package(&[
            ("word/a.xml", b"one", CompressionMethod::Stored),
            ("word/b.xml", b"two", CompressionMethod::Stored),
        ]);
        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Docx, Some(16), Some(6), None).unwrap();
        let operation = session.begin_operation("parse").unwrap();
        let first = session.open_entry_by_index(operation, 0).unwrap();
        let second = session.open_entry_by_index(operation, 1).unwrap();
        assert_eq!(drain(&mut session, operation, first, 2).unwrap(), b"one");
        assert_eq!(drain(&mut session, operation, second, 2).unwrap(), b"two");
        assert_eq!(session.usage().distinct_inflated_bytes, 6);
    }

    #[test]
    fn interleaved_readers_recompute_the_shared_total_allowance() {
        let bytes = package(&[
            ("word/a.xml", b"1234", CompressionMethod::Stored),
            ("word/b.xml", b"5678", CompressionMethod::Stored),
        ]);
        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Docx, Some(16), Some(6), None).unwrap();
        let operation = session.begin_operation("parse").unwrap();
        // Both readers are opened before either consumes the package-wide
        // budget. The second pull must observe what the first pull charged.
        let first = session.open_entry(operation, "word/a.xml").unwrap();
        let second = session.open_entry(operation, "word/b.xml").unwrap();
        assert_eq!(drain(&mut session, operation, first, 4).unwrap(), b"1234");

        let error = drain(&mut session, operation, second, 4).unwrap_err();
        assert!(error.contains("distinct-inflated-bytes"));
        assert!(error.contains("\"observed\":7"));
    }

    #[test]
    fn ordinary_crc_corruption_does_not_poison_other_operations() {
        let mut bytes = package(&[
            ("word/bad.bin", b"bad", CompressionMethod::Stored),
            ("word/good.bin", b"good", CompressionMethod::Stored),
        ]);
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        let bad_start = archive.by_index_raw(0).unwrap().data_start() as usize;
        drop(archive);
        bytes[bad_start] ^= 0xff;

        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Docx, Some(64), Some(64), None).unwrap();
        let bad_operation = session.begin_operation("bad").unwrap();
        let bad_reader = session.open_entry(bad_operation, "word/bad.bin").unwrap();
        assert!(drain(&mut session, bad_operation, bad_reader, 2)
            .unwrap_err()
            .contains("CRC"));

        let good_operation = session.begin_operation("good").unwrap();
        let good_reader = session.open_entry(good_operation, "word/good.bin").unwrap();
        assert_eq!(
            drain(&mut session, good_operation, good_reader, 2).unwrap(),
            b"good"
        );
    }

    #[test]
    fn cancel_finish_release_and_close_are_idempotent() {
        let bytes = package(&[("word/a.xml", b"1234", CompressionMethod::Stored)]);
        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Docx, Some(16), Some(16), None).unwrap();
        let operation = session.begin_operation("parse").unwrap();
        let reader = session.open_entry(operation, "word/a.xml").unwrap();
        session.release_entry(reader);
        session.release_entry(reader);
        session.cancel_operation(operation).unwrap();
        let final_usage = session.usage_for_operation(operation).unwrap();
        assert_eq!(final_usage.operation_inflated_bytes, 0);
        session.cancel_operation(operation).unwrap();
        session.finish_operation(operation).unwrap();
        session.finish_operation(operation).unwrap();
        session.close();
        session.close();
        assert!(session.begin_operation("later").is_err());
    }

    #[test]
    fn finalized_operation_history_is_bounded_without_limiting_session_lifetime() {
        let bytes = package(&[("word/a.xml", b"1", CompressionMethod::Stored)]);
        let mut session =
            PackageSession::open(bytes, OoxmlFormat::Docx, Some(16), Some(16), None).unwrap();
        let mut first = None;
        let mut latest = None;
        for index in 0..(MAX_FINALIZED_OPERATION_RECORDS + 10) {
            let operation = session
                .begin_operation(format!("operation-{index}"))
                .unwrap();
            first.get_or_insert(operation);
            session.finish_operation(operation).unwrap();
            latest = Some(operation);
        }
        assert!(session.usage_for_operation(first.unwrap()).is_none());
        assert!(session.usage_for_operation(latest.unwrap()).is_some());
        assert_eq!(session.operations.len(), MAX_FINALIZED_OPERATION_RECORDS);
    }

    #[test]
    fn handle_clones_share_one_session_and_distinct_accounting() {
        let bytes = package(&[("xl/sharedStrings.xml", b"shared", CompressionMethod::Stored)]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Xlsx, Some(64), Some(64), None).unwrap();
        let clone = handle.clone();
        assert!(Rc::ptr_eq(&handle.inner, &clone.inner));

        let mut first = handle.begin_operation("first").unwrap();
        let mut second = clone.begin_operation("second").unwrap();
        assert_eq!(first.read_bytes("xl/sharedStrings.xml").unwrap(), b"shared");
        assert_eq!(
            second.read_bytes("xl/sharedStrings.xml").unwrap(),
            b"shared"
        );
        assert_eq!(handle.usage().distinct_inflated_bytes, 6);
        assert_eq!(first.usage().unwrap().operation_inflated_bytes, 6);
        assert_eq!(second.usage().unwrap().operation_inflated_bytes, 6);
        first.finish().unwrap();
        second.finish().unwrap();
    }

    #[test]
    fn package_metadata_preserves_validated_index_order() {
        let bytes = package(&[
            ("xl/a.xml", b"first", CompressionMethod::Stored),
            ("xl/c.xml", b"other", CompressionMethod::Deflated),
            ("xl/b.xml", b"last", CompressionMethod::Stored),
        ]);
        let session =
            PackageSession::open(bytes, OoxmlFormat::Xlsx, Some(64), Some(64), None).unwrap();
        assert_eq!(session.entry_count(), 3);
        assert!(session.contains_entry("xl/a.xml"));
        assert!(!session.contains_entry("xl/missing.xml"));
        assert_eq!(
            session.entry_paths(),
            vec![
                "xl/a.xml".to_string(),
                "xl/c.xml".to_string(),
                "xl/b.xml".to_string()
            ]
        );

        let handle = PackageSessionHandle::new(session);
        assert_eq!(handle.entry_count(), 3);
        assert!(handle.contains_entry("xl/a.xml"));
        assert_eq!(handle.entry_paths()[1], "xl/c.xml");
        let operation = handle.begin_operation("lookup").unwrap();
        assert_eq!(operation.read_bytes("xl/a.xml").unwrap(), b"first");

        handle.close();
        assert_eq!(handle.entry_count(), 0);
        assert!(!handle.contains_entry("xl/a.xml"));
        assert!(handle.entry_paths().is_empty());
    }

    #[test]
    fn package_open_rejects_duplicate_and_ascii_case_colliding_item_names() {
        let mut exact = package(&[
            ("xl/a.xml", b"first", CompressionMethod::Stored),
            ("xl/b.xml", b"last", CompressionMethod::Stored),
        ]);
        // ZipWriter rejects duplicate names. Forge equal-length local and
        // central-directory names after producing an otherwise valid archive.
        for offset in 0..=exact.len() - b"xl/b.xml".len() {
            if &exact[offset..offset + b"xl/b.xml".len()] == b"xl/b.xml" {
                exact[offset..offset + b"xl/b.xml".len()].copy_from_slice(b"xl/a.xml");
            }
        }
        let exact_error = PackageSession::open(exact, OoxmlFormat::Xlsx, Some(64), Some(64), None)
            .err()
            .unwrap();
        assert!(exact_error.contains("ZIP item names must be unique"));
        assert!(!exact_error.starts_with("OOXML_RESOURCE_LIMIT:"));

        let case_collision = package(&[
            ("xl/a.xml", b"first", CompressionMethod::Stored),
            ("XL/A.XML", b"second", CompressionMethod::Stored),
        ]);
        let case_error =
            PackageSession::open(case_collision, OoxmlFormat::Xlsx, Some(64), Some(64), None)
                .err()
                .unwrap();
        assert!(case_error.contains("unique ignoring ASCII case"));
        assert!(!case_error.starts_with("OOXML_RESOURCE_LIMIT:"));
    }

    #[test]
    fn owned_entry_streams_can_be_interleaved() {
        let bytes = package(&[
            ("xl/a.xml", b"abcdef", CompressionMethod::Deflated),
            ("xl/b.xml", b"12345", CompressionMethod::Stored),
        ]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Xlsx, Some(64), Some(64), None).unwrap();
        let operation = handle.begin_operation("sheet").unwrap();
        let mut first = operation.open_entry("xl/a.xml").unwrap();
        let mut second = operation.open_entry("xl/b.xml").unwrap();
        let mut buffer = [0u8; 2];

        assert_eq!(first.read(&mut buffer).unwrap(), 2);
        assert_eq!(&buffer, b"ab");
        assert_eq!(second.read(&mut buffer).unwrap(), 2);
        assert_eq!(&buffer, b"12");
        let mut first_rest = Vec::new();
        let mut second_rest = Vec::new();
        first.read_to_end(&mut first_rest).unwrap();
        second.read_to_end(&mut second_rest).unwrap();
        assert_eq!(first_rest, b"cdef");
        assert_eq!(second_rest, b"345");
    }

    #[test]
    fn operation_finish_cancel_and_drop_are_idempotent() {
        let bytes = package(&[("word/a.xml", b"body", CompressionMethod::Stored)]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Docx, Some(64), Some(64), None).unwrap();

        let mut finished = handle.begin_operation("finished").unwrap();
        let finished_id = finished.id;
        finished.finish().unwrap();
        finished.finish().unwrap();
        finished.cancel().unwrap();
        assert!(finished.read_head("word/a.xml", 0).is_err());
        assert_eq!(
            handle.inner.borrow().operations[&finished_id].status,
            OperationStatus::Finished
        );

        let mut canceled = handle.begin_operation("canceled").unwrap();
        let canceled_id = canceled.id;
        canceled.cancel().unwrap();
        canceled.cancel().unwrap();
        canceled.finish().unwrap();
        assert!(canceled.read_head("word/a.xml", 0).is_err());
        assert_eq!(
            handle.inner.borrow().operations[&canceled_id].status,
            OperationStatus::Canceled
        );

        let dropped_id = {
            let dropped = handle.begin_operation("dropped").unwrap();
            dropped.id
        };
        assert_eq!(
            handle.inner.borrow().operations[&dropped_id].status,
            OperationStatus::Canceled
        );
    }

    #[test]
    fn entry_stream_reports_eof_after_returning_exact_credit() {
        let bytes = package(&[("ppt/slide.xml", b"12345678", CompressionMethod::Deflated)]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Pptx, Some(8), Some(8), None).unwrap();
        let operation = handle.begin_operation("slide").unwrap();
        let mut stream = operation.open_entry("ppt/slide.xml").unwrap();
        let mut exact = [0u8; 8];
        assert_eq!(stream.read(&mut exact).unwrap(), 8);
        assert_eq!(&exact, b"12345678");
        assert_eq!(stream.read(&mut exact).unwrap(), 0);
        assert_eq!(stream.read(&mut exact).unwrap(), 0);
    }

    #[test]
    fn stream_resource_failure_poisons_the_shared_handle() {
        let bytes = package(&[
            ("word/a.xml", b"1234", CompressionMethod::Stored),
            ("word/b.xml", b"5678", CompressionMethod::Stored),
        ]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Docx, Some(64), Some(6), None).unwrap();
        let operation = handle.begin_operation("parse").unwrap();
        assert_eq!(operation.read_bytes("word/a.xml").unwrap(), b"1234");
        let error = operation.read_bytes("word/b.xml").unwrap_err();
        assert!(error.contains("distinct-inflated-bytes"));
        assert_eq!(handle.assert_healthy().unwrap_err(), error);
        assert_eq!(
            operation.read_head("word/missing.xml", 0).unwrap_err(),
            error
        );
        assert!(handle.inner.borrow().readers.is_empty());
    }

    #[test]
    fn operation_limit_reporter_poison_is_shared_with_sibling_operations() {
        let bytes = package(&[(
            "xl/worksheets/sheet1.xml",
            b"<x/>",
            CompressionMethod::Stored,
        )]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Xlsx, Some(64), Some(64), None).unwrap();
        let primary = handle.begin_operation("parse-sheet").unwrap();
        let sibling = handle.begin_operation("inspect").unwrap();
        let stream = primary.open_entry("xl/worksheets/sheet1.xml").unwrap();
        let reporter = stream.limit_reporter().unwrap();

        reporter
            .observe_hard_limit(
                HardResourceLimitKind::XmlEventBytes,
                Some("xl/worksheets/sheet1.xml"),
                8,
                8,
            )
            .unwrap();
        let error = reporter
            .observe_hard_limit(
                HardResourceLimitKind::XmlEventBytes,
                Some("xl/worksheets/sheet1.xml"),
                8,
                9,
            )
            .unwrap_err();
        assert!(error.starts_with("OOXML_RESOURCE_LIMIT:"));
        assert!(error.contains("\"operation\":\"parse-sheet\""));
        assert!(error.contains("\"resource\":\"xml-event\""));
        assert!(error.contains("\"part\":\"xl/worksheets/sheet1.xml\""));
        assert_eq!(handle.assert_healthy().unwrap_err(), error);
        assert_eq!(
            sibling
                .read_head("xl/worksheets/sheet1.xml", 1)
                .unwrap_err(),
            error
        );
        assert!(handle.inner.borrow().readers.is_empty());
    }

    #[test]
    fn limit_reporter_does_not_outlive_operation_lifecycle() {
        let bytes = package(&[("word/document.xml", b"<w/>", CompressionMethod::Stored)]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Docx, Some(64), Some(64), None).unwrap();
        let reporter = {
            let mut operation = handle.begin_operation("parse").unwrap();
            let stream = operation.open_entry("word/document.xml").unwrap();
            let reporter = stream.limit_reporter().unwrap();
            operation.finish().unwrap();
            reporter
        };
        let error = reporter
            .observe_hard_limit(HardResourceLimitKind::XmlNestingDepth, None, 256, 257)
            .unwrap_err();
        assert!(error.contains("operation is not active"));
        assert!(handle.assert_healthy().is_ok());
    }

    #[test]
    fn read_head_releases_its_entry_reader() {
        let bytes = package(&[("word/a.xml", b"abcdefgh", CompressionMethod::Stored)]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Docx, Some(64), Some(64), None).unwrap();
        let operation = handle.begin_operation("inspect").unwrap();
        assert!(operation.read_head("word/missing.xml", 0).is_err());
        assert_eq!(operation.read_head("word/a.xml", 0).unwrap(), b"");
        assert_eq!(operation.usage().unwrap().operation_inflated_bytes, 0);
        assert_eq!(operation.read_head("word/a.xml", 3).unwrap(), b"abc");
        assert!(handle.inner.borrow().readers.is_empty());
        assert_eq!(operation.usage().unwrap().operation_inflated_bytes, 3);
        assert_eq!(handle.usage().distinct_inflated_bytes, 3);
    }

    #[test]
    fn active_reader_count_and_id_allocation_have_hard_stops() {
        let bytes = package(&[("word/a.xml", b"a", CompressionMethod::Stored)]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Docx, Some(64), Some(64), None).unwrap();
        let operation = handle.begin_operation("many-readers").unwrap();
        let mut streams = Vec::with_capacity(HARD_MAX_ACTIVE_ENTRY_READERS);
        for _ in 0..HARD_MAX_ACTIVE_ENTRY_READERS {
            streams.push(operation.open_entry("word/a.xml").unwrap());
        }
        let ceiling = operation.open_entry("word/a.xml").err().unwrap();
        assert!(ceiling.contains("active entry readers"));
        streams.pop();
        streams.push(operation.open_entry("word/a.xml").unwrap());
        drop(streams);
        assert!(handle.inner.borrow().readers.is_empty());

        handle.inner.borrow_mut().next_reader_id = u64::MAX;
        let exhausted = operation.open_entry("word/a.xml").err().unwrap();
        assert!(exhausted.contains("reader ID space is exhausted"));
        assert!(handle.inner.borrow().readers.is_empty());
    }

    #[test]
    fn finishing_or_dropping_an_operation_invalidates_live_streams() {
        let bytes = package(&[("word/a.xml", b"body", CompressionMethod::Stored)]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Docx, Some(64), Some(64), None).unwrap();

        let mut finished = handle.begin_operation("finished").unwrap();
        let mut finished_stream = finished.open_entry("word/a.xml").unwrap();
        finished.finish().unwrap();
        assert!(handle.inner.borrow().readers.is_empty());
        let mut byte = [0u8; 1];
        assert!(finished_stream.read(&mut byte).is_err());

        let mut dropped_stream = {
            let dropped = handle.begin_operation("dropped").unwrap();
            dropped.open_entry("word/a.xml").unwrap()
        };
        assert!(handle.inner.borrow().readers.is_empty());
        assert!(dropped_stream.read(&mut byte).is_err());
    }

    #[test]
    fn closing_a_handle_invalidates_all_clones_and_owned_streams() {
        let bytes = package(&[("word/a.xml", b"body", CompressionMethod::Stored)]);
        let handle =
            PackageSessionHandle::open(bytes, OoxmlFormat::Docx, Some(64), Some(64), None).unwrap();
        let clone = handle.clone();
        let operation = clone.begin_operation("parse").unwrap();
        let mut stream = operation.open_entry("word/a.xml").unwrap();

        handle.close();
        handle.close();
        assert!(clone.assert_healthy().is_err());
        let mut byte = [0u8; 1];
        assert_eq!(
            stream.read(&mut byte).unwrap_err().kind(),
            std::io::ErrorKind::Other
        );
    }
}
