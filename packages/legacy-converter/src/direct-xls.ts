import {
  validateLegacyXlsSourceDescriptor,
  type LegacyXlsDirectSourceDescriptor,
} from '@silurus/ooxml-core/internal/legacy-xls-source';
import wasmAssetUrl from './wasm-direct-xls/legacy_office_converter_bg.wasm?url';

export interface LegacyXlsSourceOptions { readonly wasmUrl?: string }

/** Describe the built-in direct XLS source without loading its dedicated WASM. */
export function createLegacyXlsSource(
  options: LegacyXlsSourceOptions = {},
): Readonly<LegacyXlsDirectSourceDescriptor> {
  return validateLegacyXlsSourceDescriptor({
    protocol: 'ooxml-legacy-xls-source/v1',
    builtin: 'xls',
    wasmUrl: options.wasmUrl ?? new URL(wasmAssetUrl, import.meta.url).href,
  });
}

export type { LegacyXlsDirectSourceDescriptor };
