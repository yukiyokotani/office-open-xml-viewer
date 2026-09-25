import type { ModelSource } from '@silurus/ooxml-core';
import { createLegacySource, type LegacySourceOptions } from './legacy-source.js';
import defaultWasmUrl from './wasm-direct-ppt/legacy_ppt_direct_bg.wasm?url';
import defaultModuleUrl from './legacy-ppt-source-module.ts?worker&url';

export type LegacyPptSourceOptions = LegacySourceOptions;

/**
 * A model source for legacy PowerPoint 97-2003 (.ppt) input, for
 * `LoadOptions.modelSources` of the PPTX loaders and viewers. Creating it
 * fetches nothing; the source module and its WASM load only when a .ppt input
 * is claimed.
 */
export function legacyPptSource(options: LegacyPptSourceOptions = {}): ModelSource<'pptx'> {
  return createLegacySource<'pptx'>('ppt', options, {
    wasmUrl: new URL(defaultWasmUrl, import.meta.url).href,
    moduleUrl: new URL(defaultModuleUrl, import.meta.url).href,
  });
}
