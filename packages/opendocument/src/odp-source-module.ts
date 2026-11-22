/**
 * Self-contained source module for OpenDocument presentations (.odp / .otp).
 * The PPTX renderer imports it by URL in the realm that owns the presentation
 * archive (a Worker, or Node) and calls `openModelSource`. It bundles the
 * wasm-bindgen glue and the source runtime; its WASM URL arrives in `config`,
 * the caller's ZIP resource policy in `context`.
 */
import type { ModelSourceOpenContext } from '@silurus/ooxml-core';
import {
  createOdpSourceEngine,
  resolveDirectWasmInput,
  type OdpGlue,
  type OdpNativePresentation,
} from './direct-odp-engine.js';
import { readOdfSourceModuleConfig } from './source-module-config.js';

const engine = createOdpSourceEngine(
  () => import('./wasm-odp/opendocument_odp.js') as Promise<OdpGlue>,
  resolveDirectWasmInput,
);

export async function openModelSource(
  bytes: Uint8Array,
  config: unknown,
  signal?: AbortSignal,
  _transfer?: readonly Transferable[],
  context?: ModelSourceOpenContext,
): Promise<Readonly<{ archive: OdpNativePresentation; close(): void }>> {
  const { wasmUrl } = readOdfSourceModuleConfig(config, 'OpenDocument ODP');
  const owned = await engine.open(bytes, wasmUrl, context, signal);
  return Object.freeze({ archive: owned.archive, close: () => owned.closeArchive() });
}
