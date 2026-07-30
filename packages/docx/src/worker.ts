import init, { DocxArchive, reinit } from './wasm/docx_parser.js';
import { decodeDataUrl, WasmParserHost } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types';

// RB6: a `panic = "abort"` build traps (not unwinds) on a Rust panic / OOM /
// stack overflow, poisoning this worker's single WASM instance so every LATER
// file would crash on the corrupted memory too. `WasmParserHost` draws the line
// between a graceful `Result::Err` (instance stays healthy) and a trap (instance
// recycled): `host.run(...)` catches a trap, frees the archive, marks the
// instance poisoned, and `host.ensureReady()` respawns a fresh module before the
// next request — so one bad file fails alone and the next parses on clean memory.
//
// The host also OWNS the archive handle (`host.archive`): a
// `DocxArchive(bytes, max)` copies the file into WASM ONCE and scans the central
// directory ONCE, then a later `extractImage` reads media by zip path straight
// from the retained archive. Freed + replaced on a re-parse, and freed + nulled
// by the host itself on a trap so a later parse never double-frees a handle from
// a discarded instance.
const host = new WasmParserHost<DocxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.type === 'init') {
    host.setWasmUrl(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  // Echo the correlation id so the client routes the response to the right
  // pending promise (id correlation, not response-type matching).
  const id = req.id;
  try {
    await host.ensureReady();
    if (req.type === 'parse') {
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
      const bytes = new Uint8Array(req.data);
      // Both the construction and `parse()` run under `host.run` so a trap in
      // EITHER poisons + recycles the instance (and frees the archive). Adopting
      // via `setArchive` frees any prior handle first — the re-parse dispose.
      // `parse()` returns the model as UTF-8 JSON bytes on success and throws a
      // JS Error on parse/serialize failure (Result<Vec<u8>, JsValue>). The throw
      // is caught by the outer try/catch below. wasm-bindgen hands back a fresh
      // Uint8Array (a copy of the Rust Vec), so its buffer is exclusively ours:
      // forward it to the main thread as a transferable — no clone, no decode
      // here. The single decode + JSON.parse happens once, on the main thread.
      const json = host.run(() => {
        const archive = new DocxArchive(bytes, max, maxTotal, maxEntries);
        host.setArchive(archive);
        return archive.parse();
      });
      const documentJson = json.buffer as ArrayBuffer;
      const res: WorkerResponse = { type: 'parsed', id, documentJson };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [
        documentJson,
      ]);
      return;
    }

    const archive = host.archive;

    if (req.type === 'extractImage') {
      if (!archive) throw new Error('No docx loaded');
      // wasm-bindgen already hands back a fresh, standalone Uint8Array here (its
      // glue does `getArrayU8FromWasm0(ptr,len).slice()` then frees the Rust Vec),
      // so `.buffer` is a full-span, non-WASM-backed ArrayBuffer we own outright —
      // transfer it directly. A second `new Uint8Array(bytes).slice()` would just
      // re-copy the whole entry for nothing.
      const out = host.run(() => archive.extract_image(req.path).buffer as ArrayBuffer);
      const res: WorkerResponse = { type: 'imageExtracted', id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [out]);
      return;
    }
    if (req.type === 'toMarkdown') {
      if (!archive) throw new Error('No docx loaded');
      // Project the already-opened handle to markdown (no re-copy of the file,
      // no re-scan of the central directory). A plain string has no transferable
      // backing, so it is posted by structured clone like any other value.
      const markdown = host.run(() => archive.to_markdown());
      const res: WorkerResponse = { type: 'markdownRendered', id, markdown };
      self.postMessage(res);
      return;
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', id, message: String(err) };
    self.postMessage(res);
  }
};
