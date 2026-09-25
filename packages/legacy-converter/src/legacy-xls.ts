import type { ModelSource } from '@silurus/ooxml-core';
import { createLegacySource, type LegacySourceOptions } from './legacy-source.js';
import defaultWasmUrl from './wasm-direct-xls/legacy_xls_direct_bg.wasm?url';
import defaultModuleUrl from './legacy-xls-source-module.ts?worker&url';

export type LegacyXlsSourceOptions = LegacySourceOptions;

/**
 * A model source for legacy Excel 97-2003 (.xls) input, for
 * `LoadOptions.modelSources` of the XLSX loaders and viewers. Creating it
 * fetches nothing; the source module and its WASM load only when a .xls input
 * is claimed.
 */
export function legacyXlsSource(options: LegacyXlsSourceOptions = {}): ModelSource<'xlsx'> {
  return createLegacySource<'xlsx'>('xls', options, {
    wasmUrl: new URL(defaultWasmUrl, import.meta.url).href,
    moduleUrl: new URL(defaultModuleUrl, import.meta.url).href,
  });
}
