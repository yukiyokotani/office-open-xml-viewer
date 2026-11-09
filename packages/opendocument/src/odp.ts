import { resolveBundledAssetUrl } from '@silurus/ooxml-core/internal/bundled-asset-url';
import type { ModelSource } from '@silurus/ooxml-core';
import { createOdfSource, deferredOdfSource, type OpenDocumentSourceOptions } from './odf-source.js';
import defaultWasmUrl from './wasm-odp/opendocument_odp_bg.wasm?url';
import defaultModuleUrl from './odp-source-module.ts?url';

export type OdpOptions = OpenDocumentSourceOptions;

/**
 * Create a model source for OpenDocument presentations (.odp / .otp) with
 * custom asset URLs, for example when an asset pipeline relocates the bundled
 * files. Pass it as the `odp` option of the PPTX loaders and viewers.
 * Creating it fetches nothing; the source module and its WASM load only when
 * an ODP input is claimed.
 */
export function createOdp(options: OdpOptions = {}): ModelSource<'pptx'> {
  if (typeof options !== 'object' || options === null || Array.isArray(options))
    throw new TypeError('OpenDocument ODP source options must be an object');
  return createOdfSource<'pptx'>('odp', options, {
    wasmUrl: resolveBundledAssetUrl(defaultWasmUrl, import.meta.url),
    moduleUrl: resolveBundledAssetUrl(defaultModuleUrl, import.meta.url),
  });
}

/**
 * OpenDocument presentation reader with the bundled asset URLs:
 * `import { odp } from '@silurus/ooxml-opendocument/odp'`, then pass `{ odp }`
 * to the PPTX loaders and viewers. Importing it fetches and instantiates
 * nothing; the bundled URLs resolve on first use, and the source module and
 * its WASM load only when an ODP input is claimed.
 */
export const odp: ModelSource<'pptx'> = /* @__PURE__ */ deferredOdfSource('pptx', () => createOdp());
