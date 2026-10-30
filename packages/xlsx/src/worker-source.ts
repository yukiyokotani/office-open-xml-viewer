import {
  normalizeXlsxWorksheetPolicy,
  xlsxWorksheetPolicyForWasm,
  DEFAULT_XLSX_WORKSHEET_POLICY,
  type NormalizedXlsxWorksheetPolicy,
} from './internal/worksheet-policy.js';
import init, { XlsxArchive, reinit } from './wasm/xlsx_parser.js';
import {
  decodeDataUrl,
  WasmParserHost,
} from '@silurus/ooxml-core';
import {
  decodeOoxmlResourceUsage,
  resourcePolicyForWasm,
  serializeWorkerError,
  type PullSessionCommand,
} from '@silurus/ooxml-core/worker';
import type { WorkerRequest, WorkerResponse } from './types.js';
import { readXlsxArchiveBootstrap } from './internal/archive-bootstrap-source.js';
import { isWorksheetPullCommand, WorksheetPullWorker } from './worksheet-pull-source-worker.js';
import type { WorkerWorksheetSourceOwner } from './internal/worker-worksheet-source.js';

// RB6: a `panic = "abort"` build traps (not unwinds) on a Rust panic / OOM /
// stack overflow, poisoning this worker's single WASM instance so every LATER
// file (or sheet) would crash on the corrupted memory too. `WasmParserHost`
// draws the line between a graceful `Result::Err` (instance stays healthy) and a
// trap (instance recycled): `host.run(...)` catches a trap, frees the archive,
// marks the instance poisoned, and `host.ensureReady()` respawns a fresh module
// before the next request — so one bad file fails alone and the next parses on
// clean memory.
//
// The host also OWNS the archive handle (`host.archive`): a
// `XlsxArchive(bytes, max)` copies the file into WASM ONCE and scans the central
// directory ONCE; workbook/sharedStrings/theme state is reused by the bounded
// worksheet pull sessions. `extractImage` also reads by zip path
// straight from the retained archive. Freed + replaced on a re-parse, and freed +
// nulled by the host itself on a trap so a later parse never double-frees a
// handle from a discarded instance.
const host = new WasmParserHost<XlsxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});
let source: WorkerWorksheetSourceOwner<XlsxArchive> | undefined;
// Per-document worksheet policy, captured when a parse claims a new document.
let worksheetPolicy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY;
const worksheetPull = new WorksheetPullWorker(
  () => source?.cursor() ?? host.archive,
  undefined,
  (operation) => {
    if (source) return source.execute(operation);
    const archive = host.archive;
    if (!archive) throw new Error('Workbook not loaded');
    return host.run(() => operation(archive));
  },
  undefined,
  undefined,
  () => worksheetPolicy,
);

self.onmessage = async (e: MessageEvent<WorkerRequest | PullSessionCommand<number>>) => {
  const req = e.data;

  // A model source's host-layout reply is consumed by its own listener
  // (requestHostLayoutFromPage) and never enters the request-id dispatcher.
  if (source?.isHostLayoutResult(req)) return;

  if (isWorksheetPullCommand(req)) {
    await worksheetPull.dispatchSafely(req, (response, transfer) =>
      (self.postMessage as (message: unknown, transfer?: Transferable[]) => void)(response, transfer),
    );
    return;
  }

  if (req.type === 'init') {
    host.setWasmInput(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  // Every non-init request carries a correlation id that must be echoed back so
  // the client can route the response to the right pending promise.
  const id = req.id;
  if (req.type === 'openSheetSession') worksheetPull.reserveOpen(req);
  try {
    // Validate the request policy before any reset / instance / source work.
    // Older requests without a policy resolve to the defaults.
    const parsePolicy = req.type === 'parse'
      ? normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: req.worksheetPolicy?.worksheet })
      : undefined;
    if (req.type === 'openSheetSession') {
      if (!source) await host.ensureReady();
      if (source?.cursor()) source.execute((archive) => archive.assert_healthy());
      else if (host.archive) host.run(() => host.archive?.assert_healthy());
      await worksheetPull.open(req.sheetIndex, req.sheetName, req);
      await worksheetPull.postOpenedSafely(
        req,
        () => self.postMessage({
          type: 'sheetSessionOpened',
          id,
          sessionId: req.sessionId,
          operationId: req.operationId,
          generation: req.generation,
        } satisfies WorkerResponse),
        (error) => self.postMessage({
          type: 'error',
          id,
          ...serializeWorkerError(error),
        } satisfies WorkerResponse),
      );
      return;
    }
    if (req.type === 'parse') await worksheetPull.reset();
    await (req.type === 'extractImage' ? worksheetPull.runDuringPull : worksheetPull.run).call(worksheetPull, async () => {
    if (req.type === 'parse' ? !req.source : !source) await host.ensureReady();
    if (req.type !== 'parse' && (source?.cursor() ?? host.archive)) {
      if (source) source.execute((archive) => archive.assert_healthy());
      else host.run(() => host.archive?.assert_healthy());
    }
    if (req.type === 'parse') {
      source?.closeModelSource();
      source = undefined;
      worksheetPolicy = parsePolicy ?? DEFAULT_XLSX_WORKSHEET_POLICY;
      if (req.source) {
        host.disposeArchive();
        if (!req.sourceOwnerUrl) throw new TypeError('XLSX source owner URL is missing');
        const { WorkerWorksheetSourceOwner } = await import(/* @vite-ignore */ req.sourceOwnerUrl) as typeof import('./internal/worker-worksheet-source.js');
        const owner = new WorkerWorksheetSourceOwner(host);
        source = owner;
        // The renderer lives on the page in this mode, so the page measures.
        await owner.openModelSource(
          new Uint8Array(req.data),
          req.source,
          (font) => owner.requestHostLayoutFromPage(self as unknown as Parameters<typeof owner.requestHostLayoutFromPage>[0], font),
          req.sourceTransfer,
          { resourcePolicy: req.resourcePolicy, worksheetPolicy },
        );
      } else {
        const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(req.resourcePolicy);
        const bytes = new Uint8Array(req.data);
        host.run(() => {
          const opened = new XlsxArchive(bytes, maxEntry, maxTotal, maxEntries);
          // Adopt first so the host owns cleanup if the setter throws/traps.
          host.setArchive(opened);
          opened.set_worksheet_limits(...xlsxWorksheetPolicyForWasm(worksheetPolicy));
          return opened;
        });
      }
      // Both the construction and `parse()` run under `host.run` so a trap in
      // EITHER poisons + recycles the instance (and frees the archive). Adopting
      // via `setArchive` frees any prior handle first — the re-parse dispose.
      // `parse()` returns the workbook index as UTF-8 JSON bytes (Result<Vec<u8>,
      // JsValue>). wasm-bindgen hands back a fresh Uint8Array that owns its
      // buffer, so forward it to main as a transferable — no clone, no decode
      // here. The single decode + JSON.parse happens on main.
      const { workbook: json, usage } = readXlsxArchiveBootstrap(
        () => source
          ? source.execute((current) => current.parse())
          : host.run(() => host.archive!.parse()),
        () => source
          ? source.resourceUsage()
          : host.run(() => host.archive!.resource_usage()),
      );
      const workbookJson = source
        ? source.copyBytes(json)
        : json.buffer as ArrayBuffer;
      const maximumDigitWidth = source?.maximumDigitWidth;
      const res: WorkerResponse = {
        type: 'parsed', id, workbookJson, usage,
        ...(maximumDigitWidth === undefined ? {} : { layoutMetrics: { maximumDigitWidth } }),
      };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [
        workbookJson,
      ]);
      return;
    }

    const archive = source?.cursor() ?? host.archive;

    if (req.type === 'extractImage') {
      if (!archive) throw new Error('No xlsx loaded');
      // wasm-bindgen already hands back a fresh, standalone Uint8Array here (its
      // glue does `getArrayU8FromWasm0(ptr,len).slice()` then frees the Rust Vec),
      // so `.buffer` is a full-span, non-WASM-backed ArrayBuffer we own outright —
      // transfer it directly. A second `new Uint8Array(bytes).slice()` would just
      // re-copy the whole entry for nothing.
      const out = source
        ? source.extractImage(req.path)
        : host.run(() => archive.extract_image(req.path).buffer as ArrayBuffer);
      const res: WorkerResponse = { type: 'imageExtracted', id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [out]);
      return;
    }

    if (req.type === 'resourceUsage') {
      if (!archive) throw new Error('No xlsx loaded');
      const bytes = source ? source.resourceUsage() : host.run(() => host.archive!.resource_usage());
      const usage = bytes === undefined ? undefined : decodeOoxmlResourceUsage(bytes);
      self.postMessage({ type: 'resourceUsage', id, usage } satisfies WorkerResponse);
      return;
    }

    if (req.type === 'toMarkdown') {
      if (!archive) throw new Error('No xlsx loaded');
      // Project the already-opened handle to markdown (no re-copy of the file,
      // no re-scan of the central directory). A plain string has no transferable
      // backing, so it is posted by structured clone like any other value.
      const markdown = source ? source.toMarkdown() : host.run(() => host.archive!.to_markdown());
      const res: WorkerResponse = { type: 'markdownRendered', id, markdown };
      self.postMessage(res);
      return;
    }
    });
  } catch (err) {
    if (req.type === 'openSheetSession') worksheetPull.abandonOpen(req.sessionId);
    if (req.type === 'parse') {
      try { source?.closeModelSource(); } catch {}
    }
    const res: WorkerResponse = { type: 'error', id, ...serializeWorkerError(err) };
    try {
      self.postMessage(res);
    } catch {
      // A parser/operation error was preserved in `res`, but the response
      // channel itself is unavailable. Never reject the async event handler.
    }
  }
};
