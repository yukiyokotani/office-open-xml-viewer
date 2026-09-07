import {
  validateLegacyPptSourceDescriptor,
  type LegacyPptDirectSourceDescriptor,
} from '@silurus/ooxml-core/internal/legacy-ppt-source';
import wasmAssetUrl from './wasm-direct-ppt/legacy_office_converter_bg.wasm?url';

export interface LegacyPptSourceOptions {
  /** Override the emitted dedicated direct-PPT WASM asset URL. */
  readonly wasmUrl?: string;
}

/**
 * Describe the built-in direct PPT source without fetching or initializing it.
 * The consumer that owns the presentation session resolves this descriptor.
 */
export function createLegacyPptSource(
  options: LegacyPptSourceOptions = {},
): Readonly<LegacyPptDirectSourceDescriptor> {
  const wasmUrl = options.wasmUrl ?? new URL(wasmAssetUrl, import.meta.url).href;
  return validateLegacyPptSourceDescriptor({
    protocol: 'ooxml-legacy-ppt-source/v1',
    builtin: 'ppt',
    wasmUrl,
  });
}

export type { LegacyPptDirectSourceDescriptor };
