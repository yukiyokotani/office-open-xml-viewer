/**
 * Self-contained source module for OpenDocument text (.odt / .ott). The DOCX
 * renderer imports it by URL in the realm that owns the document archive (a
 * Worker, or Node) and calls `openModelSource`. It bundles the wasm-bindgen
 * glue and the source runtime; its WASM URL arrives in `config`, the caller's
 * ZIP resource policy in `context`.
 */
import type { ModelSourceOpenContext } from '@silurus/ooxml-core';
import {
  createOdtSourceEngine,
  resolveDirectWasmInput,
  type OdtGlue,
  type OdtNativeDocument,
} from './direct-odt-engine.js';
import { readOdfSourceModuleConfig } from './source-module-config.js';

const engine = createOdtSourceEngine(
  () => import('./wasm-odt/opendocument_odt.js') as Promise<OdtGlue>,
  resolveDirectWasmInput,
);

export async function openModelSource(
  bytes: Uint8Array,
  config: unknown,
  signal?: AbortSignal,
  _transfer?: readonly Transferable[],
  context?: ModelSourceOpenContext,
): Promise<Readonly<{ archive: OdtNativeDocument; close(): void }>> {
  const { wasmUrl } = readOdfSourceModuleConfig(config, 'OpenDocument ODT');
  const owned = await engine.open(bytes, wasmUrl, context, signal);
  return Object.freeze({ archive: owned.archive, close: () => owned.closeArchive() });
}
