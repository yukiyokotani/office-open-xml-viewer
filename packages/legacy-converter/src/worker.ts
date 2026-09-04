import {
  installLegacyOfficeConversionWorkerHandler,
  type LegacyOfficeWorkerRequest,
  type LegacyOfficeConversionWorkerScope,
} from '@silurus/ooxml-core';
import type { InitInput } from './wasm/legacy_office_converter.js';
import { createLegacyOfficeWasmConverterFromSource } from './engine.js';

type ConfiguredWorkerRequest = LegacyOfficeWorkerRequest & {
  readonly converterWasmUrl?: unknown;
};

let resolveWasm: ((source: InitInput) => void) | undefined;
let rejectWasm: ((error: Error) => void) | undefined;
const wasmSource = new Promise<InitInput>((resolve, reject) => {
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
    resolveWasm?.(new URL(value));
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

installLegacyOfficeConversionWorkerHandler(
  workerScope,
  createLegacyOfficeWasmConverterFromSource(() => wasmSource),
);
