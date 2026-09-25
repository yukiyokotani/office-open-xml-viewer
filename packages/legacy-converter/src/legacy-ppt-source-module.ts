/**
 * Self-contained source module for legacy PowerPoint 97-2003 (.ppt) input.
 * The PPTX renderer imports it by URL in the realm that owns the presentation
 * archive (a Worker, or Node) and calls `openModelSource`. It bundles the
 * wasm-bindgen glue and the direct reader runtime; its WASM URL arrives in
 * `config`.
 */
import {
  createLegacyPptSourceEngine,
  resolveDirectWasmInput,
  type LegacyPptGlue,
  type LegacyPptNativeArchive,
} from './direct-ppt-engine.js';
import { readLegacySourceModuleConfig } from './source-module-config.js';

const engine = createLegacyPptSourceEngine(
  () => import('./wasm-direct-ppt/legacy_ppt_direct.js') as Promise<LegacyPptGlue>,
  resolveDirectWasmInput,
);

export async function openModelSource(
  bytes: Uint8Array,
  config: unknown,
  signal?: AbortSignal,
): Promise<Readonly<{ archive: LegacyPptNativeArchive; close(): void }>> {
  const { wasmUrl, maxInputBytes } = readLegacySourceModuleConfig(config, 'legacy PPT');
  if (bytes.byteLength > maxInputBytes) {
    throw new RangeError('legacy PPT input exceeds the configured size limit');
  }
  const owned = await engine.open(bytes, wasmUrl, signal);
  return Object.freeze({ archive: owned.archive, close: () => owned.closeArchive() });
}
