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
import {
  acquirePptxSessionFromArchive,
  type PptxNodeAcquisition,
  type PptxNodeAcquisitionOptions,
  type PptxNodeArchive,
} from './node-session-acquisition.js';
// @ts-ignore wasm-pack generated module has no declaration entry
import * as pptxWasm from '../wasm/pptx_parser.js';

export {
  acquirePptxSessionFromArchive,
  type PptxNodeAcquisition,
  type PptxNodeAcquisitionOptions,
  type PptxNodeArchive,
} from './node-session-acquisition.js';

interface PptxArchiveConstructor {
  new (
    data: Uint8Array,
    maxArchiveEntryBytes?: bigint | null,
    maxTotalInflatedBytes?: bigint | null,
    maxArchiveEntries?: bigint | null,
  ): PptxNodeArchive;
}

let runtimeModule: WebAssembly.Module | undefined;
let runtimeHost: WasmRuntimeGenerationHost<PptxNodeArchive> | undefined;

function formatRuntime(module: WebAssembly.Module): WasmRuntimeGenerationHost<PptxNodeArchive> {
  if (!runtimeHost) {
    runtimeModule = module;
    runtimeHost = new WasmRuntimeGenerationHost(
      pptxWasm as unknown as WasmModuleRuntime,
      module,
    );
  } else if (runtimeModule !== module) {
    throw new Error('PPTX runtime was already initialized with another WebAssembly.Module');
  }
  return runtimeHost;
}

/** Format-owned archive acquisition and bootstrap projection for Node. */
export async function acquirePptxNodeSession(
  bytes: Uint8Array,
  module: WebAssembly.Module,
  options: PptxNodeAcquisitionOptions = {},
): Promise<PptxNodeAcquisition> {
  const resourceOptions = normalizeLoadResourceOptions(options);
  const metrics = new OoxmlResourceMetricsSession({
    enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
    format: 'pptx',
    mode: 'node',
    scope: 'session',
    policy: resourceOptions.policy,
    onMetrics: resourceOptions.onResourceMetrics,
    emitToConsole: resourceOptions.debug,
  });
  metrics.setSourceBytes(bytes.byteLength);
  let handle: WasmArchiveHandle<PptxNodeArchive> | undefined;
  let admissionOwnsFailure = false;
  try {
    throwIfAborted(options.signal);
    const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(resourceOptions.policy);
    const Archive = (pptxWasm as unknown as { PptxArchive: PptxArchiveConstructor }).PptxArchive;
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
    admissionOwnsFailure = true;
    return acquirePptxSessionFromArchive({
      archive,
      sourceByteLength: bytes.byteLength,
      closeArchive: () => handle?.close((current: PptxNodeArchive) => current.free()),
    }, options, metrics);
  } catch (error) {
    try { handle?.close((archive: PptxNodeArchive) => archive.free()); } catch {}
    const normalized = parseResourceLimitError(error) ?? error;
    if (!admissionOwnsFailure) metrics.fail(normalized);
    throw normalized;
  }
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  throw createAbortError();
}

function createAbortError(): Error {
  const error = new Error('PPTX presentation session was aborted');
  error.name = 'AbortError';
  return error;
}
