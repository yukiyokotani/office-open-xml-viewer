import { resolveBundledAssetUrl } from '@silurus/ooxml-core/internal/bundled-asset-url';
import type { ModelSource } from '@silurus/ooxml-core';
import { createOdfSource, deferredOdfSource, type OpenDocumentSourceOptions } from './odf-source.js';
import defaultWasmUrl from './wasm-odt/opendocument_odt_bg.wasm?url';
import defaultModuleUrl from './odt-source-module.ts?url';

export type OdtOptions = OpenDocumentSourceOptions;

/**
 * Create a model source for OpenDocument text (.odt / .ott) with custom asset
 * URLs, for example when an asset pipeline relocates the bundled files. Pass
 * it as the `odt` option of the DOCX loaders and viewers. Creating it fetches
 * nothing; the source module and its WASM load only when an ODT input is
 * claimed.
 */
export function createOdt(options: OdtOptions = {}): ModelSource<'docx'> {
  if (typeof options !== 'object' || options === null || Array.isArray(options))
    throw new TypeError('OpenDocument ODT source options must be an object');
  return createOdfSource<'docx'>('odt', options, {
    wasmUrl: resolveBundledAssetUrl(defaultWasmUrl, import.meta.url),
    moduleUrl: resolveBundledAssetUrl(defaultModuleUrl, import.meta.url),
  });
}

/**
 * OpenDocument text reader with the bundled asset URLs:
 * `import { odt } from '@silurus/ooxml-opendocument/odt'`, then pass `{ odt }`
 * to the DOCX loaders and viewers. Importing it fetches and instantiates
 * nothing; the bundled URLs resolve on first use, and the source module and its
 * WASM load only when an ODT input is claimed.
 */
export const odt: ModelSource<'docx'> = /* @__PURE__ */ deferredOdfSource('docx', () => createOdt());
