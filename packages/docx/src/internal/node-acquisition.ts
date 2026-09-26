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
  WasmRuntimeGenerationHost,
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

/** An already-opened archive admitted into the Node DOCX session. */
export interface DocxOwnedArchiveSource {
  readonly archive: DocxNodeSessionArchive;
  readonly sourceByteLength: number;
  closeArchive(): void;
}

interface DocxArchiveConstructor {
  new (
    data: Uint8Array,
    maxArchiveEntryBytes?: bigint | null,
    maxTotalInflatedBytes?: bigint | null,
    maxArchiveEntries?: bigint | null,
  ): DocxNodeArchive;
}

let runtimeModule: WebAssembly.Module | undefined;
let runtimeHost: WasmRuntimeGenerationHost<DocxNodeArchive> | undefined;

function formatRuntime(module: WebAssembly.Module): WasmRuntimeGenerationHost<DocxNodeArchive> {
  if (!runtimeHost) {
    runtimeModule = module;
    runtimeHost = new WasmRuntimeGenerationHost(
      docxWasm as unknown as WasmModuleRuntime,
      module,
    );
  } else if (runtimeModule !== module) {
    throw new Error('DOCX runtime was already initialized with another WebAssembly.Module');
  }
  return runtimeHost;
}

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
  module: WebAssembly.Module,
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
  let admissionOwnsFailure = false;
  try {
    throwIfAborted(options.signal);
    const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(resourceOptions.policy);
    const Archive = (docxWasm as unknown as { DocxArchive: DocxArchiveConstructor }).DocxArchive;
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
    metrics.checkpoint('container ready');
    const openedHandle = handle;
    admissionOwnsFailure = true;
    return await consumeDocxArchive(
      archive,
      () => openedHandle.close((current: DocxNodeArchive) => current.free()),
      metrics,
      options,
      consume,
    );
  } catch (error) {
    if (admissionOwnsFailure) throw error;
    try { handle?.close((archive: DocxNodeArchive) => archive.free()); } catch {}
    const normalized = parseTypedParserError(error) ?? error;
    metrics.fail(normalized);
    throw normalized;
  }
}

/**
 * Admit an already-opened archive (for example one returned by an
 * application-supplied model source) into the same acknowledged body cursor,
 * accounting and cleanup as an OOXML package. The session owns `closeArchive`
 * from the moment this is called, including on failure.
 */
export async function acquireDocxSessionFromArchive<TResult>(
  source: DocxOwnedArchiveSource,
  options: DocxNodeAcquisitionOptions,
  consume: (
    transport: DocxNodePullTransport,
    identity: DocxNodePullIdentity,
    options: DocxNodePullOptions,
  ) => Promise<TResult>,
): Promise<AcquiredDocxNodeDocument<TResult>> {
  let closed = false;
  const closeArchive = (): void => {
    if (closed) return;
    closed = true;
    source.closeArchive();
  };
  let metrics: OoxmlResourceMetricsSession | undefined;
  try {
    if (!Number.isSafeInteger(source.sourceByteLength) || source.sourceByteLength < 0) {
      throw new RangeError('DOCX sourceByteLength must be a non-negative safe integer');
    }
    const resourceOptions = normalizeLoadResourceOptions(options);
    metrics = new OoxmlResourceMetricsSession({
      enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
      format: 'docx',
      mode: 'node',
      scope: 'session',
      policy: resourceOptions.policy,
      onMetrics: resourceOptions.onResourceMetrics,
      emitToConsole: resourceOptions.debug,
    });
    metrics.setSourceBytes(source.sourceByteLength);
    throwIfAborted(options.signal);
    metrics.checkpoint('container ready');
  } catch (error) {
    try { closeArchive(); } catch {}
    const normalized = parseTypedParserError(error) ?? error;
    metrics?.fail(normalized);
    throw normalized;
  }
  return consumeDocxArchive(source.archive, closeArchive, metrics, options, consume);
}

async function consumeDocxArchive<TResult>(
  archive: DocxNodeSessionArchive,
  closeArchive: () => void,
  metrics: OoxmlResourceMetricsSession,
  options: DocxNodeAcquisitionOptions,
  consume: (
    transport: DocxNodePullTransport,
    identity: DocxNodePullIdentity,
    options: DocxNodePullOptions,
  ) => Promise<TResult>,
): Promise<AcquiredDocxNodeDocument<TResult>> {
  let pull: DocumentPullWorker | undefined;
  let transport: DocxNodePullTransport | undefined;
  try {
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
    return { archive, result, usage, metrics, closeArchive };
  } catch (error) {
    await pull?.reset().catch(() => undefined);
    transport?.terminate();
    try { closeArchive(); } catch {}
    const normalized = parseTypedParserError(error) ?? error;
    metrics.fail(normalized);
    throw normalized;
  }
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  throw createAbortError();
}

function createAbortError(): Error {
  const error = new Error('DOCX document session was aborted');
  error.name = 'AbortError';
  return error;
}
