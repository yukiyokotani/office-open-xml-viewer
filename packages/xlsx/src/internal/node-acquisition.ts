import {
  normalizeXlsxWorksheetPolicy,
  xlsxWorksheetPolicyForWasm,
  type NormalizedXlsxWorksheetPolicy,
} from './worksheet-policy.js';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  parseTypedParserError,
  resourcePolicyForWasm,
} from '@silurus/ooxml-core/worker';
import {
  wasmRuntimeSingleton,
  type WasmArchiveHandle,
  type WasmModuleRuntime,
} from '@silurus/ooxml-core/internal/wasm-runtime-generation';
import type { ParsedWorkbook } from '../types.js';
import { readXlsxArchiveBootstrap } from './archive-bootstrap.js';
// @ts-ignore wasm-pack generated module has no declaration entry
import * as xlsxWasm from '../wasm/xlsx_parser.js';
import { createAbortError, throwIfAborted } from '@silurus/ooxml-core/internal/abort';

const ABORT_MESSAGE = 'XLSX workbook session was aborted';

export interface XlsxNodeAcquisitionOptions {
  readonly resourceLimits?: import('@silurus/ooxml-core').OoxmlResourceLimits;
  readonly maxZipEntryBytes?: number;
  readonly debug?: boolean;
  readonly onResourceMetrics?: (metrics: import('@silurus/ooxml-core').OoxmlResourceMetrics) => void;
  readonly signal?: AbortSignal;
  readonly xlsxWorksheetLimits?: Readonly<import('@silurus/ooxml-core').XlsxWorksheetLimits>;
}

export interface XlsxNodeArchive {
  free(): void;
  set_worksheet_limits(
    maxRows: bigint,
    maxCells: bigint,
    maxOwnedUtf8Bytes: bigint,
    maxJsonBytes: bigint,
  ): void;
  parse(): Uint8Array;
  resource_usage(): Uint8Array;
  /** Raw bytes of one picture part (or a model source's picture key). */
  extract_image(path: string): Uint8Array;
  /** OOXML packages only: whether the part's content type is `image/*`. A
   * model-source archive omits it because its picture keys are already the
   * only readable entries. */
  is_picture_part?(path: string): boolean;
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

const formatRuntime = wasmRuntimeSingleton<XlsxNodeArchive>(
  xlsxWasm as unknown as WasmModuleRuntime,
  'XLSX',
);

export interface XlsxNodeAcquisition {
  readonly archive: XlsxNodeSessionArchive;
  readonly workbookIndex: ParsedWorkbook;
  readonly usage: OoxmlResourceUsageSnapshot | undefined;
  readonly metrics: OoxmlResourceMetricsSession;
  /** Normalized, deeply frozen worksheet policy captured at acquisition. */
  readonly worksheetPolicy: NormalizedXlsxWorksheetPolicy;
  closeArchive(): void;
}

/**
 * The archive a Node XLSX session reads after acquisition: the XLSX parser
 * archive, or a model-source archive whose ZIP accounting is optional.
 */
export interface XlsxNodeSessionArchive
  extends Omit<
    XlsxNodeArchive,
    'free' | 'resource_usage' | 'sheet_cursor_resource_usage' | 'set_worksheet_limits'
  > {
  /** Absent when the source has no ZIP accounting; absence is not zero usage. */
  resource_usage?(): Uint8Array;
  sheet_cursor_resource_usage?(): Uint8Array;
}

/** Format-owned archive acquisition and workbook-index projection for Node. */
export async function acquireXlsxNodeSession(
  bytes: Uint8Array,
  wasmModule: WebAssembly.Module,
  options: XlsxNodeAcquisitionOptions = {},
): Promise<XlsxNodeAcquisition> {
  const worksheetPolicy = normalizeXlsxWorksheetPolicy(options);
  const resourceOptions = normalizeLoadResourceOptions(options);
  const metrics = new OoxmlResourceMetricsSession({
    enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
    format: 'xlsx',
    mode: 'node',
    scope: 'session',
    policy: resourceOptions.policy,
    formatLimits: { xlsxWorksheetLimits: { ...worksheetPolicy.worksheet } },
    onMetrics: resourceOptions.onResourceMetrics,
    emitToConsole: resourceOptions.debug,
  });
  metrics.setSourceBytes(bytes.byteLength);
  let handle: WasmArchiveHandle<XlsxNodeArchive> | undefined;
  try {
    throwIfAborted(options.signal, ABORT_MESSAGE);
    const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(resourceOptions.policy);
    const Archive = (xlsxWasm as unknown as { XlsxArchive: XlsxArchiveConstructor }).XlsxArchive;
    handle = await formatRuntime(wasmModule).open(
      () => new Archive(bytes, maxEntry, maxTotal, maxEntries),
      {
        signal: options.signal,
        abortError: () => createAbortError(ABORT_MESSAGE),
        disposeOnAbort: (archive) => archive.free(),
      },
    );
    throwIfAborted(options.signal, ABORT_MESSAGE);
    const archive = handle.proxy;
    const [maxRows, maxCells, maxOwnedUtf8Bytes, maxJsonBytes] =
      xlsxWorksheetPolicyForWasm(worksheetPolicy);
    archive.set_worksheet_limits(maxRows, maxCells, maxOwnedUtf8Bytes, maxJsonBytes);
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
      worksheetPolicy,
      closeArchive: () => handle?.close((current: XlsxNodeArchive) => current.free()),
    };
  } catch (error) {
    try { handle?.close((archive: XlsxNodeArchive) => archive.free()); } catch {}
    const normalized = parseTypedParserError(error) ?? error;
    metrics.fail(normalized);
    throw normalized;
  }
}
