import type { ModelSource } from '@silurus/ooxml-core';
import { createLegacySource, type LegacySourceOptions } from './legacy-source.js';
import defaultWasmUrl from './wasm-direct-doc/legacy_doc_direct_bg.wasm?url';
import defaultModuleUrl from './legacy-doc-source-module.ts?worker&url';

export type LegacyDocSourceOptions = LegacySourceOptions;

/**
 * A model source for legacy Word 97-2003 (.doc) input, for
 * `LoadOptions.modelSources` of the DOCX loaders and viewers. Creating it
 * fetches nothing; the source module and its WASM load only when a .doc input
 * is claimed.
 */
export function legacyDocSource(options: LegacyDocSourceOptions = {}): ModelSource<'docx'> {
  return createLegacySource<'docx'>('doc', options, {
    wasmUrl: new URL(defaultWasmUrl, import.meta.url).href,
    moduleUrl: new URL(defaultModuleUrl, import.meta.url).href,
  });
}
