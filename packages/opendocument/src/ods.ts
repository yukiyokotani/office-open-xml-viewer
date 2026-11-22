import { resolveBundledAssetUrl } from '@silurus/ooxml-core/internal/bundled-asset-url';
import type { ModelSource } from '@silurus/ooxml-core';
import { createOdfSource, deferredOdfSource, type OpenDocumentSourceOptions } from './odf-source.js';
import defaultWasmUrl from './wasm-ods/opendocument_ods_bg.wasm?url';
import defaultModuleUrl from './ods-source-module.ts?url';

export type OdsOptions = OpenDocumentSourceOptions;

/**
 * Create a model source for OpenDocument spreadsheets (.ods / .ots) with
 * custom asset URLs, for example when an asset pipeline relocates the bundled
 * files. Pass it as the `ods` option of the XLSX loaders and viewers.
 * Creating it fetches nothing; the source module and its WASM load only when
 * an ODS input is claimed.
 */
export function createOds(options: OdsOptions = {}): ModelSource<'xlsx'> {
  if (typeof options !== 'object' || options === null || Array.isArray(options))
    throw new TypeError('OpenDocument ODS source options must be an object');
  return createOdfSource<'xlsx'>('ods', options, {
    wasmUrl: resolveBundledAssetUrl(defaultWasmUrl, import.meta.url),
    moduleUrl: resolveBundledAssetUrl(defaultModuleUrl, import.meta.url),
  });
}

/**
 * OpenDocument spreadsheet reader with the bundled asset URLs:
 * `import { ods } from '@silurus/ooxml-opendocument/ods'`, then pass `{ ods }`
 * to the XLSX loaders and viewers (`new XlsxViewer(container, { ods })`).
 * Importing it fetches and instantiates nothing; the bundled URLs resolve on
 * first use, and the source module and its WASM load only when an ODS input
 * is claimed.
 */
export const ods: ModelSource<'xlsx'> = /* @__PURE__ */ deferredOdfSource('xlsx', () => createOds());
