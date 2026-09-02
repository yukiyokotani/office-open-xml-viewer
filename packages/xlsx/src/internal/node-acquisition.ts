import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  parseResourceLimitError,
  resourcePolicyForWasm,
} from '@silurus/ooxml-core/worker';
import {
  WasmRuntimeGenerationHost,
  type WasmArchiveHandle,
  type WasmModuleRuntime,
} from '@silurus/ooxml-core/internal/wasm-runtime-generation';
import type { ParsedWorkbook } from '../types.js';
import { readXlsxArchiveBootstrap } from './archive-bootstrap.js';
// @ts-ignore wasm-pack generated module has no declaration entry
import * as xlsxWasm from '../wasm/xlsx_parser.js';

export interface XlsxNodeAcquisitionOptions {
  readonly resourceLimits?: import('@silurus/ooxml-core').OoxmlResourceLimits;
  readonly maxZipEntryBytes?: number;
  readonly debug?: boolean;
  readonly onResourceMetrics?: (metrics: import('@silurus/ooxml-core').OoxmlResourceMetrics) => void;
  readonly signal?: AbortSignal;
}

export interface XlsxNodeArchive {
  free(): void;
  parse(): Uint8Array;
  resource_usage(): Uint8Array;
  open_sheet_cursor(sheetIndex: number, name: string): void;
  pull_sheet_cursor(rowCredit: number): Uint8Array;
  sheet_cursor_pull_finished(): boolean;
  sheet_cursor_resource_usage(): Uint8Array;
  acknowledge_sheet_cursor_terminal(): void;
  cancel_sheet_cursor(): void;
  close_sheet_cursor(): void;
}

interface XlsxArchiveConstructor {
  new (
    data: Uint8Array,
    maxArchiveEntryBytes?: bigint | null,
    maxTotalInflatedBytes?: bigint | null,
    maxArchiveEntries?: bigint | null,
  ): XlsxNodeArchive;
}

let runtimeModule: WebAssembly.Module | undefined;
let runtimeHost: WasmRuntimeGenerationHost<XlsxNodeArchive> | undefined;

function formatRuntime(module: WebAssembly.Module): WasmRuntimeGenerationHost<XlsxNodeArchive> {
  if (!runtimeHost) {
    runtimeModule = module;
    runtimeHost = new WasmRuntimeGenerationHost(
      xlsxWasm as unknown as WasmModuleRuntime,
      module,
    );
  } else if (runtimeModule !== module) {
    throw new Error('XLSX runtime was already initialized with another WebAssembly.Module');
  }
  return runtimeHost;
}

export interface XlsxNodeAcquisition {
  readonly archive: XlsxNodeArchive;
  readonly workbookIndex: ParsedWorkbook;
  readonly usage: OoxmlResourceUsageSnapshot | undefined;
  readonly metrics: OoxmlResourceMetricsSession;
  closeArchive(): void;
}

/** Format-owned archive acquisition and workbook-index projection for Node. */
export async function acquireXlsxNodeSession(
  bytes: Uint8Array,
  module: WebAssembly.Module,
  options: XlsxNodeAcquisitionOptions = {},
): Promise<XlsxNodeAcquisition> {
  const resourceOptions = normalizeLoadResourceOptions(options);
  const metrics = new OoxmlResourceMetricsSession({
    enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
    format: 'xlsx',
    mode: 'node',
    scope: 'session',
    policy: resourceOptions.policy,
    onMetrics: resourceOptions.onResourceMetrics,
    emitToConsole: resourceOptions.debug,
  });
  metrics.setSourceBytes(bytes.byteLength);
  let handle: WasmArchiveHandle<XlsxNodeArchive> | undefined;
  try {
    throwIfAborted(options.signal);
    const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(resourceOptions.policy);
    const Archive = (xlsxWasm as unknown as { XlsxArchive: XlsxArchiveConstructor }).XlsxArchive;
    handle = await formatRuntime(module).open(
      () => new Archive(bytes, maxEntry, maxTotal, maxEntries),
      {
        signal: options.signal,
        abortError: createAbortError,
        disposeOnAbort: (archive) => archive.free(),
      },
    );
    throwIfAborted(options.signal);
    const archive = handle.proxy;
    const { workbook: workbookIndex, usage } = readXlsxArchiveBootstrap(
      () => JSON.parse(new TextDecoder().decode(archive.parse())) as ParsedWorkbook,
      () => archive.resource_usage(),
    );
    metrics.observeUsage(usage);
    metrics.checkpoint('workbook index ready');
    return {
      archive,
      workbookIndex,
      usage,
      metrics,
      closeArchive: () => handle?.close((current: XlsxNodeArchive) => current.free()),
    };
  } catch (error) {
    try { handle?.close((archive: XlsxNodeArchive) => archive.free()); } catch {}
    const normalized = parseResourceLimitError(error) ?? error;
    metrics.fail(normalized);
    throw normalized;
  }
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  throw createAbortError();
}

function createAbortError(): Error {
  const error = new Error('XLSX workbook session was aborted');
  error.name = 'AbortError';
  return error;
}
