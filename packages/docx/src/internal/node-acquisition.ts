import { normalizeDocxParserError } from './parser-errors.js';
import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  parseTypedParserError,
  resourcePolicyForWasm,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import { InProcessPullTransport } from '@silurus/ooxml-core/internal/in-process-pull-transport';
import {
  wasmRuntimeSingleton,
  type WasmArchiveHandle,
  type WasmModuleRuntime,
} from '@silurus/ooxml-core/internal/wasm-runtime-generation';
import type { DocxDocumentCursorArchive } from '../document-pull-worker.js';
import {
  DocumentPullWorker,
  readDocxDocumentCursorUsage,
} from '../document-pull-worker.js';
// @ts-ignore wasm-pack generated module has no declaration entry
import * as docxWasm from '../wasm/docx_parser.js';
import { createAbortError, throwIfAborted } from '@silurus/ooxml-core/internal/abort';

const ABORT_MESSAGE = 'DOCX document session was aborted';

export interface DocxNodeAcquisitionOptions {
  readonly resourceLimits?: import('@silurus/ooxml-core').OoxmlResourceLimits;
  readonly maxZipEntryBytes?: number;
  readonly debug?: boolean;
  readonly onResourceMetrics?: (metrics: import('@silurus/ooxml-core').OoxmlResourceMetrics) => void;
  readonly signal?: AbortSignal;
}

export interface DocxNodeArchive extends DocxDocumentCursorArchive {
  free(): void;
  extract_image(path: string): Uint8Array;
  /** OOXML packages only: whether the part's content type is `image/*`. A
   * model-source archive omits it because its picture keys are already the
   * only readable entries. */
  is_picture_part?(path: string): boolean;
  /** `undefined` when the package has no document-cursor checkpoint. */
  document_cursor_resource_usage(): Uint8Array | undefined;
  resource_usage(): Uint8Array;
}

/**
 * The archive a Node DOCX session reads after acquisition: the DOCX parser
 * archive, or a model-source archive whose ZIP accounting is optional.
 */
export interface DocxNodeSessionArchive extends DocxDocumentCursorArchive {
  extract_image(path: string): Uint8Array;
  /** Absent when the source has no ZIP accounting; absence is not zero usage. */
  resource_usage?(): Uint8Array;
}


interface DocxArchiveConstructor {
  new (
    data: Uint8Array,
    maxArchiveEntryBytes?: bigint | null,
    maxTotalInflatedBytes?: bigint | null,
    maxArchiveEntries?: bigint | null,
  ): DocxNodeArchive;
}

const formatRuntime = wasmRuntimeSingleton<DocxNodeArchive>(
  docxWasm as unknown as WasmModuleRuntime,
  'DOCX',
);

export type DocxNodePullIdentity = Readonly<{
  sessionId: number;
  operationId: number;
  generation: number;
}>;
export type DocxNodePullTransport = InProcessPullTransport<
  PullSessionResponse<ArrayBuffer, number>
>;
export type DocxNodePullOptions = Readonly<{
  signal?: AbortSignal;
  onUsage: (usage: OoxmlResourceUsageSnapshot) => void;
}>;

export interface AcquiredDocxNodeDocument<TResult> {
  readonly archive: DocxNodeSessionArchive;
  readonly result: TResult;
  readonly usage: OoxmlResourceUsageSnapshot | undefined;
  readonly metrics: OoxmlResourceMetricsSession;
  closeArchive(): void;
}

/** Format-owned DOCX archive, cursor transport, accounting, and cleanup. */
export async function acquireDocxNodeDocument<TResult>(
  bytes: Uint8Array,
  wasmModule: WebAssembly.Module,
  options: DocxNodeAcquisitionOptions,
  consume: (
    transport: DocxNodePullTransport,
    identity: DocxNodePullIdentity,
    options: DocxNodePullOptions,
  ) => Promise<TResult>,
): Promise<AcquiredDocxNodeDocument<TResult>> {
  const resourceOptions = normalizeLoadResourceOptions(options);
  const metrics = new OoxmlResourceMetricsSession({
    enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
    format: 'docx',
    mode: 'node',
    scope: 'session',
    policy: resourceOptions.policy,
    onMetrics: resourceOptions.onResourceMetrics,
    emitToConsole: resourceOptions.debug,
  });
  metrics.setSourceBytes(bytes.byteLength);
  let handle: WasmArchiveHandle<DocxNodeArchive> | undefined;
  let pull: DocumentPullWorker | undefined;
  let transport: DocxNodePullTransport | undefined;
  try {
    throwIfAborted(options.signal, ABORT_MESSAGE);
    const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(resourceOptions.policy);
    const Archive = (docxWasm as unknown as { DocxArchive: DocxArchiveConstructor }).DocxArchive;
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
    metrics.checkpoint('container ready');
    pull = new DocumentPullWorker(() => archive);
    const identity = { sessionId: 1, operationId: 1, generation: 1 } as const;
    pull.open(identity);
    transport = new InProcessPullTransport(
      (command, respond) => pull?.dispatch(command as PullSessionCommand<number>, respond),
      () => undefined,
    );
    let usage: OoxmlResourceUsageSnapshot | undefined;
    const result = await consume(transport, identity, {
      signal: options.signal,
      onUsage: (checkpoint) => {
        usage = checkpoint;
        metrics.observeUsage(checkpoint);
      },
    });
    usage ??= readDocxDocumentCursorUsage((operation) => operation(archive));
    metrics.observeUsage(usage);
    metrics.checkpoint('model streamed');
    await pull.reset();
    transport.terminate();
    return {
      archive,
      result,
      usage,
      metrics,
      closeArchive: () => handle?.close((current: DocxNodeArchive) => current.free()),
    };
  } catch (error) {
    await pull?.reset().catch(() => undefined);
    transport?.terminate();
    try { handle?.close((archive: DocxNodeArchive) => archive.free()); } catch {}
    const normalized = parseTypedParserError(error) ?? normalizeDocxParserError(error);
    metrics.fail(normalized);
    throw normalized;
  }
}
