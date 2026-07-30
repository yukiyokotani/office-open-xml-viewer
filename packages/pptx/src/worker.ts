import { decodeDataUrl, WasmParserHost } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types';
import init, { PptxArchive, reinit } from './wasm/pptx_parser.js';

// RB6: a `panic = "abort"` build traps (not unwinds) on a Rust panic / OOM /
// stack overflow, poisoning this worker's single WASM instance so every LATER
// file would crash on the corrupted memory too. `WasmParserHost` draws the line
// between a graceful `Result::Err` (instance stays healthy) and a trap (instance
// recycled): `host.run(...)` catches a trap, frees the archive, marks the
// instance poisoned, and `host.ensureReady()` respawns a fresh module before the
// next request — so one bad file fails alone and the next parses on clean memory.
//
// The host also OWNS the archive handle (`host.archive`): a
// `PptxArchive(bytes, max)` copies the file into WASM ONCE and scans the central
// directory ONCE, then a later `extractMedia` / `extractImage` reads by zip path
// straight from the retained archive (no re-copy, no re-open, no JS-side buffer
// kept alive — the sole copy lives in WASM). Freed + replaced on a re-parse, and
// freed + nulled by the host itself on a trap so a later parse never
// double-frees a handle from a discarded instance.
const host = new WasmParserHost<PptxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery: `init` re-run is a no-op against wasm-bindgen's cached
  // singleton, so a trap would poison every LATER file. `reinit` nulls the
  // singleton first, forcing a genuine re-instantiation on fresh linear memory.
  reinit,
});

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.kind === 'init') {
    // Retain the init lifecycle in the host rather than a `ready` flag +
    // handshake. Every request below `await`s `ensureReady()`, so a REJECTED
    // init rejects the request (the catch posts an `error` response the bridge
    // turns into a rejected `load()`), never a silent hang on a main-side `ready`
    // wait. After a trap, `ensureReady()` respawns a fresh module here.
    host.setWasmUrl(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  // Echo the correlation id so the client routes the response to the right
  // pending promise (id correlation, not response-type matching).
  const id = req.id;
  try {
    await host.ensureReady();
    if (req.kind === 'parse') {
      const max =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      const maxTotal =
        typeof req.maxZipTotalBytes === 'number' && req.maxZipTotalBytes > 0
          ? BigInt(req.maxZipTotalBytes)
          : undefined;
      const maxEntries =
        typeof req.maxZipEntries === 'number' && req.maxZipEntries > 0
          ? BigInt(req.maxZipEntries)
          : undefined;
      const bytes = new Uint8Array(req.buffer);
      // Both the construction and `parse()` run under `host.run` so a trap in
      // EITHER poisons + recycles the instance (and frees the archive). Adopting
      // via `setArchive` frees any prior handle first — the re-parse dispose.
      // `parse()` returns the model as UTF-8 JSON bytes (Result<Vec<u8>,
      // JsValue>). wasm-bindgen hands back a fresh Uint8Array that owns its
      // buffer, so forward it to the main thread as a transferable — no clone,
      // no decode here. The single decode + JSON.parse happens once, on main.
      const json = host.run(() => {
        const archive = new PptxArchive(bytes, max, maxTotal, maxEntries);
        host.setArchive(archive);
        return archive.parse();
      });
      const presentationJson = json.buffer as ArrayBuffer;
      const msg: WorkerResponse = { kind: 'parsed', id, presentationJson };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [
        presentationJson,
      ]);
      return;
    }

    const archive = host.archive;

    if (req.kind === 'extractMedia') {
      if (!archive) throw new Error('No pptx loaded');
      // wasm-bindgen already hands back a fresh, standalone Uint8Array here (its
      // glue does `getArrayU8FromWasm0(ptr,len).slice()` then frees the Rust Vec),
      // so `bytes.buffer` is a full-span, non-WASM-backed ArrayBuffer we own
      // outright — transfer it directly. A second `new Uint8Array(bytes).slice()`
      // would just re-copy the whole entry for nothing.
      const out = host.run(() => archive.extract_media(req.path).buffer as ArrayBuffer);
      const msg: WorkerResponse = { kind: 'mediaExtracted', id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [out]);
      return;
    }

    if (req.kind === 'extractImage') {
      if (!archive) throw new Error('No pptx loaded');
      // See extractMedia above: the extracted Uint8Array already owns a
      // standalone full-span buffer, so transfer it without a second copy.
      const out = host.run(() => archive.extract_image(req.path).buffer as ArrayBuffer);
      const msg: WorkerResponse = { kind: 'imageExtracted', id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [out]);
      return;
    }

    if (req.kind === 'toMarkdown') {
      if (!archive) throw new Error('No pptx loaded');
      // Project the already-opened handle to markdown (no re-copy of the file,
      // no re-scan of the central directory). A plain string has no transferable
      // backing, so it is posted by structured clone like any other value.
      const markdown = host.run(() => archive.to_markdown());
      const msg: WorkerResponse = { kind: 'markdownRendered', id, markdown };
      self.postMessage(msg);
      return;
    }
  } catch (err) {
    const msg: WorkerResponse = {
      kind: 'error',
      id,
      message: err instanceof Error ? err.message : String(err),
    };
    self.postMessage(msg);
  }
};
