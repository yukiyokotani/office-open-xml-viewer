/**
 * Self-contained source module for OpenDocument spreadsheets (.ods / .ots).
 * The XLSX renderer imports it by URL in the realm that owns the workbook
 * archive (a Worker, or Node) and calls `openModelSource`. It bundles the
 * wasm-bindgen glue and the source runtime; its WASM URL arrives in
 * `config`, the caller's ZIP resource policy in `context`. The archive's
 * host-layout capability lets the renderer measure the default cell font,
 * whose digit width encodes the authored column widths.
 */
import type { ModelSourceOpenContext } from '@silurus/ooxml-core';
import {
  createOdsSourceEngine,
  resolveDirectWasmInput,
  type OdsGlue,
  type OdsNativeArchive,
} from './direct-ods-engine.js';
import { readOdfSourceModuleConfig } from './source-module-config.js';

const engine = createOdsSourceEngine(
  () => import('./wasm-ods/opendocument_ods.js') as Promise<OdsGlue>,
  resolveDirectWasmInput,
);

export async function openModelSource(
  bytes: Uint8Array,
  config: unknown,
  signal?: AbortSignal,
  _transfer?: readonly Transferable[],
  context?: ModelSourceOpenContext,
): Promise<Readonly<{ archive: OdsNativeArchive; close(): void }>> {
  const { wasmUrl } = readOdfSourceModuleConfig(config, 'OpenDocument ODS');
  const owned = await engine.open(bytes, wasmUrl, context, signal);
  return Object.freeze({ archive: owned.archive, close: () => owned.closeArchive() });
}
