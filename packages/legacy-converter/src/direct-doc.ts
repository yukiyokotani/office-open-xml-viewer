import {
  validateLegacyDocSourceDescriptor,
  type LegacyDocDirectSourceDescriptor,
} from '@silurus/ooxml-core/internal/legacy-doc-source';
import wasmAssetUrl from './wasm-direct-doc/legacy_office_converter_bg.wasm?url';

export interface LegacyDocSourceOptions {
  /** Override the emitted dedicated direct-DOC WASM asset URL. */
  readonly wasmUrl?: string;
}

/** Describe the built-in direct DOC source without loading or initializing its WASM runtime. */
export function createLegacyDocSource(
  options: LegacyDocSourceOptions = {},
): Readonly<LegacyDocDirectSourceDescriptor> {
  return validateLegacyDocSourceDescriptor({
    protocol: 'ooxml-legacy-doc-source/v1',
    builtin: 'doc',
    wasmUrl: options.wasmUrl ?? new URL(wasmAssetUrl, import.meta.url).href,
  });
}

export type { LegacyDocDirectSourceDescriptor };
