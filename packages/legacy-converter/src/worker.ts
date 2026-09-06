import {
  installLegacyOfficeConversionWorkerHandler,
  type LegacyOfficeWorkerRequest,
  type LegacyOfficeConversionWorkerScope,
} from '@silurus/ooxml-core';
import type { InitInput } from './wasm/legacy_office_converter.js';
import { createLegacyOfficeWasmConverterFromSource } from './engine.js';
import { requestXlsFontMeasurement } from './xls-font-worker.js';

type ConfiguredWorkerRequest = LegacyOfficeWorkerRequest & {
  readonly converterWasmUrl?: unknown;
  readonly measureXlsNormalFont?: unknown;
};

let resolveWasm: ((source: { wasm: InitInput; measure: boolean }) => void) | undefined;
let rejectWasm: ((error: Error) => void) | undefined;
const wasmSource = new Promise<{ wasm: InitInput; measure: boolean }>((resolve, reject) => {
  resolveWasm = resolve;
  rejectWasm = reject;
});

globalThis.addEventListener('message', ((event: MessageEvent<ConfiguredWorkerRequest>) => {
  const value = event.data?.converterWasmUrl;
  if (typeof value !== 'string') {
    rejectWasm?.(new Error('missing converter WASM URL'));
    return;
  }
  try {
    resolveWasm?.({ wasm: new URL(value), measure: event.data.measureXlsNormalFont === true });
  } catch {
    rejectWasm?.(new Error('invalid converter WASM URL'));
  }
  resolveWasm = undefined;
  rejectWasm = undefined;
}) as EventListener, { once: true });

const workerScope: LegacyOfficeConversionWorkerScope = {
  postMessage(message, transfer) {
    globalThis.postMessage(message, { transfer });
  },
  addEventListener(type, listener) {
    globalThis.addEventListener(type, listener as EventListener);
  },
  removeEventListener(type, listener) {
    globalThis.removeEventListener(type, listener as EventListener);
  },
};

let converter: ReturnType<typeof createLegacyOfficeWasmConverterFromSource> | undefined;
installLegacyOfficeConversionWorkerHandler(workerScope, {
  async convert(input) {
    const config = await wasmSource;
    converter ??= createLegacyOfficeWasmConverterFromSource(() => config.wasm,
      config.measure ? requestXlsFontMeasurement(globalThis) : undefined);
    return converter.convert(input);
  },
});
