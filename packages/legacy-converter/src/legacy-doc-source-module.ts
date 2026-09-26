/**
 * Self-contained source module for legacy Word 97-2003 (.doc) input. The
 * DOCX renderer imports it by URL in the realm that owns the document archive
 * (a Worker, or Node) and calls `openModelSource`. It bundles the wasm-bindgen
 * glue and the direct reader runtime; its WASM URL arrives in `config`.
 */
import {
  createLegacyDocSourceEngine,
  legacyDocViewDefaults,
  resolveDirectWasmInput,
  type LegacyDocGlue,
  type LegacyDocNativeDocument,
} from './direct-doc-engine.js';
import { readLegacySourceModuleConfig } from './source-module-config.js';

const engine = createLegacyDocSourceEngine(
  () => import('./wasm-direct-doc/legacy_doc_direct.js') as Promise<LegacyDocGlue>,
  resolveDirectWasmInput,
);

export async function openModelSource(
  bytes: Uint8Array,
  config: unknown,
  signal?: AbortSignal,
): Promise<Readonly<{
  archive: LegacyDocNativeDocument;
  viewDefaults: Readonly<{ showTrackedChanges?: boolean }>;
  close(): void;
}>> {
  const { wasmUrl, maxInputBytes } = readLegacySourceModuleConfig(config, 'legacy DOC');
  if (bytes.byteLength > maxInputBytes) {
    throw new RangeError('legacy DOC input exceeds the configured size limit');
  }
  const owned = await engine.open(bytes, wasmUrl, signal);
  try {
    return Object.freeze({
      archive: owned.archive,
      viewDefaults: legacyDocViewDefaults(owned.archive),
      close: () => owned.closeArchive(),
    });
  } catch (error) {
    try { owned.closeArchive(); } catch {}
    throw error;
  }
}
