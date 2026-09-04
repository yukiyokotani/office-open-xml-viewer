/** Optional browser Worker transport and converter-boundary utilities. */
import LegacyOfficeWasmWorker from '../packages/legacy-converter/src/worker.ts?worker&inline';
import legacyOfficeWasmAssetUrl from '../packages/legacy-converter/src/wasm/legacy_office_converter_bg.wasm?url';
import {
  createLegacyOfficeWasmConverter,
  LEGACY_OFFICE_WASM_ENGINE,
  LEGACY_OFFICE_WASM_ENGINE_VERSION,
  type LegacyOfficeWasmConverterOptions,
} from '../packages/legacy-converter/src/index.js';
import {
  createDisposableWorkerLegacyOfficeConverter,
  type LegacyOfficeConversionWorkerAdapterOptions,
  type LegacyOfficeConversionWorker,
  type LegacyOfficeWorkerRequest,
  type LegacyOfficeConverter,
} from '../packages/core/src/index.js';

/**
 * Create the built-in converter in a disposable browser Worker. Importing this
 * opt-in entry is the only normal path that bundles the converter Worker/WASM.
 */
export function createLegacyOfficeWasmWorkerConverter(
  options: LegacyOfficeConversionWorkerAdapterOptions = {},
): LegacyOfficeConverter {
  const wasmUrl = new URL(legacyOfficeWasmAssetUrl, import.meta.url).href;
  return createDisposableWorkerLegacyOfficeConverter(
    () => configuredLegacyOfficeWorker(new LegacyOfficeWasmWorker(), wasmUrl),
    options,
  );
}

function configuredLegacyOfficeWorker(
  worker: Worker,
  converterWasmUrl: string,
): LegacyOfficeConversionWorker {
  return {
    postMessage(message: LegacyOfficeWorkerRequest, transfer: Transferable[]) {
      worker.postMessage({ ...message, converterWasmUrl }, transfer);
    },
    addEventListener(type: string, listener: EventListener) {
      worker.addEventListener(type, listener);
    },
    removeEventListener(type: string, listener: EventListener) {
      worker.removeEventListener(type, listener);
    },
    terminate() {
      worker.terminate();
    },
  };
}

export {
  createLegacyOfficeWasmConverter,
  LEGACY_OFFICE_WASM_ENGINE,
  LEGACY_OFFICE_WASM_ENGINE_VERSION,
  type LegacyOfficeWasmConverterOptions,
};

export {
  createDisposableWorkerLegacyOfficeConverter,
  installLegacyOfficeConversionWorkerHandler,
  validateConvertedOoxml,
  LegacyOfficeConversionError,
  type LegacyOfficeConversionFailureReason,
  type LegacyOfficeConversionInput,
  type LegacyOfficeConversionOptions,
  type LegacyOfficeConversionRecord,
  type LegacyOfficeConversionResult,
  type LegacyOfficeConverter,
  type LegacyOfficeFormatConversionOptions,
  type LegacyOfficeFormat,
  type LegacyOfficeConversionWorker,
  type LegacyOfficeConversionWorkerAdapterOptions,
  type LegacyOfficeConversionWorkerFactory,
  type LegacyOfficeConversionWorkerScope,
  type LegacyOfficeWorkerRequest,
  type LegacyOfficeWorkerResponse,
} from '../packages/core/src/index.js';
