import {
  type OoxmlFormat,
} from '@silurus/ooxml-core';
import {
  bindLegacyOfficeConversionSignal,
  resolveOfficeInputWithOptionalConversion,
} from '@silurus/ooxml-core/internal/legacy-office-conversion';
import type { OoxmlNodeSessionOptions } from './session-options.ts';

/** Normalize the Node facade before lazily resolving any format parser WASM. */
export async function normalizeNodeOfficeInput(
  buffer: ArrayBuffer | Uint8Array,
  format: OoxmlFormat,
  options: OoxmlNodeSessionOptions,
): Promise<Uint8Array> {
  const bound = bindLegacyOfficeConversionSignal(options.legacyConversion, format, options.signal);
  try {
    return await resolveOfficeInputWithOptionalConversion(
      buffer,
      format,
      bound.options,
      options.password,
    );
  } finally {
    bound.cleanup();
  }
}
