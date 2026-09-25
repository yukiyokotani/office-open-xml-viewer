/**
 * Self-contained source module for legacy Excel 97-2003 (.xls) input. The
 * XLSX renderer imports it by URL in the realm that owns the workbook archive
 * (a Worker, or Node) and calls `openModelSource`. It bundles the wasm-bindgen
 * glue and the direct reader runtime; its WASM URL arrives in `config`. The
 * archive's host-layout capability lets the renderer measure the Normal font.
 */
import {
  createLegacyXlsSourceEngine,
  resolveDirectWasmInput,
  type LegacyXlsGlue,
  type LegacyXlsNativeArchive,
} from './direct-xls-engine.js';
import { readLegacySourceModuleConfig } from './source-module-config.js';

const engine = createLegacyXlsSourceEngine(
  () => import('./wasm-direct-xls/legacy_xls_direct.js') as Promise<LegacyXlsGlue>,
  resolveDirectWasmInput,
);

export async function openModelSource(
  bytes: Uint8Array,
  config: unknown,
  signal?: AbortSignal,
): Promise<Readonly<{ archive: LegacyXlsNativeArchive; close(): void }>> {
  const { wasmUrl, maxInputBytes } = readLegacySourceModuleConfig(config, 'legacy XLS');
  if (bytes.byteLength > maxInputBytes) {
    throw new RangeError('legacy XLS input exceeds the configured size limit');
  }
  const owned = await engine.open(bytes, wasmUrl, signal);
  return Object.freeze({ archive: owned.archive, close: () => owned.closeArchive() });
}
