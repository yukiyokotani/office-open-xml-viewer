import init, { DocxArchive, reinit } from './wasm/docx_parser.js';
import {
  decodeDataUrl,
  WasmParserHost,
} from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  PULL_SESSION_PROTOCOL,
  resourcePolicyForWasm,
  serializeWorkerError,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import type { WorkerRequest, WorkerResponse } from './types';
import { DocumentPullWorker, isDocumentPullCommand } from './document-pull-worker.js';

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
const documentPull = new DocumentPullWorker(
  () => host.archive,
  (operation) => host.run(() => {
    const archive = host.archive;
    if (!archive) throw new Error('No docx loaded');
    return operation(archive);
  }),
);
let documentGeneration = 0;

const post = (
  message: WorkerResponse | PullSessionResponse<ArrayBuffer, number>,
  transfer?: Transferable[],
) => (self.postMessage as (message: unknown, transfer?: Transferable[]) => void)(message, transfer);

self.onmessage = async (e: MessageEvent<WorkerRequest | PullSessionCommand<number>>) => {
  const req = e.data;

  if (isDocumentPullCommand(req)) {
    try {
      await documentPull.dispatch(req, post);
    } catch (error) {
      post({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'error',
        sessionId: req.sessionId,
        operationId: req.operationId,
        generation: req.generation,
        requestId: req.requestId,
        error: serializeWorkerError(error),
      });
    }
    return;
  }

  if (req.type === 'init') {
    host.setWasmInput(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  // Echo the correlation id so the client routes the response to the right
  // pending promise (id correlation, not response-type matching).
  const id = req.id;
  try {
    await host.ensureReady();
    if (req.type !== 'parse' && host.archive) {
      const retained = host.archive;
      host.run(() => retained.assert_healthy());
    }
    if (req.type === 'parse') {
      await documentPull.reset();
      const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(req.resourcePolicy);
      const bytes = new Uint8Array(req.data);
      // Construction and every later cursor call run under `host.run`, so a
      // trap poisons and recycles the instance. The parse response opens a
      // correlated pull session; complete body units, never a monolithic model
      // JSON value, cross to Window and require consumer ACK.
      host.run(() => {
        const archive = new DocxArchive(bytes, maxEntry, maxTotal, maxEntries);
        host.setArchive(archive);
      });
      documentGeneration += 1;
      const identity = {
        sessionId: documentGeneration,
        operationId: documentGeneration,
        generation: documentGeneration,
      };
      documentPull.open(identity);
      post({ type: 'documentSessionOpened', id, ...identity });
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
    if (req.type === 'resourceUsage') {
      if (!archive) throw new Error('No docx loaded');
      const usage = decodeOoxmlResourceUsage(host.run(() => archive.resource_usage()));
      post({ type: 'resourceUsage', id, usage });
      return;
    }
    if (req.type === 'toMarkdown') {
      if (!archive) throw new Error('No docx loaded');
      // Project the already-opened handle to markdown (no re-copy of the file,
      // no re-scan of the central directory). A plain string has no transferable
      // backing, so it is posted by structured clone like any other value.
      const markdown = host.run(() => archive.to_markdown());
      const res: WorkerResponse = { type: 'markdownRendered', id, markdown };
      post(res);
      return;
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', id, ...serializeWorkerError(err) };
    post(res);
  }
};
