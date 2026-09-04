import type { InitInput } from './wasm/legacy_office_converter.js';
import wasmAssetUrl from './wasm/legacy_office_converter_bg.wasm?url';
import type { LegacyOfficeConverter } from '@silurus/ooxml-core';
import {
  createLegacyOfficeWasmConverterFromSource,
  LEGACY_OFFICE_WASM_ENGINE,
  LEGACY_OFFICE_WASM_ENGINE_VERSION,
} from './engine.js';

export { LEGACY_OFFICE_WASM_ENGINE, LEGACY_OFFICE_WASM_ENGINE_VERSION };

export interface LegacyOfficeWasmConverterOptions {
  /**
   * Override the emitted WASM asset. Supplying bytes or a compiled module is
   * useful in Node and in bundlers that do not rewrite `new URL()` assets.
   */
  readonly wasm?: InitInput;
}

/**
 * Create the purpose-built local WASM converter. In browsers this object should
 * normally be installed inside the disposable Worker adapter; direct use runs
 * synchronous conversion work on the calling realm once initialization ends.
 */
export function createLegacyOfficeWasmConverter(
  options: LegacyOfficeWasmConverterOptions = {},
): LegacyOfficeConverter {
  return createLegacyOfficeWasmConverterFromSource(() => resolveWasm(options.wasm));
}

async function resolveWasm(override: InitInput | undefined): Promise<InitInput> {
  if (override !== undefined) {
    return override;
  }
  const url = new URL(wasmAssetUrl, import.meta.url);
  if (url.protocol === 'file:' && typeof process !== 'undefined' && process.versions?.node) {
    const { readFile } = await import('node:fs/promises');
    return readFile(url);
  }
  return url;
}
