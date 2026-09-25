import {
  bindLegacyXlsHostServices,
  validateLegacyXlsSourceDescriptor,
  type LegacyXlsDirectSourceDescriptor,
} from '@silurus/ooxml-core/internal/legacy-xls-source';
import { attachXlsFontMeasurement, resolveXlsFontMeasurement } from './xls-font-worker.js';
import wasmAssetUrl from './wasm-direct-xls/legacy_office_converter_bg.wasm?url';

export interface LegacyXlsSourceOptions { readonly wasmUrl?: string }

/** Describe the built-in direct XLS source without loading its dedicated WASM. */
export function createLegacyXlsSource(
  options: LegacyXlsSourceOptions = {},
): Readonly<LegacyXlsDirectSourceDescriptor> {
  const descriptor = validateLegacyXlsSourceDescriptor({
    protocol: 'ooxml-legacy-xls-source/v1',
    builtin: 'xls',
    wasmUrl: options.wasmUrl ?? new URL(wasmAssetUrl, import.meta.url).href,
  });
  // The source owns the Normal-font measurement policy (the caller's
  // measurement, else the named font in the current document); the
  // spreadsheet host only calls these services.
  return bindLegacyXlsHostServices(descriptor, {
    resolve: resolveXlsFontMeasurement,
    attach: attachXlsFontMeasurement,
  });
}

export type { LegacyXlsDirectSourceDescriptor };
