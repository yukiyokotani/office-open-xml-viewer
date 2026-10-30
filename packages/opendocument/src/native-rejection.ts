import type { OdfFamily } from './odf-source.js';

const REASON = /^[a-z0-9][a-z0-9-]{0,63}$/;

/**
 * Decode a native reader failure. Only native Result failures enter here; WASM
 * initialization and network errors bypass it. `ODF_UNSUPPORTED:<reason>`
 * marks content outside the admitted subset (the reason is a fixed,
 * content-free code from the reader's admission tables); encryption and
 * resource limits have their own markers. Native messages never reach the v2
 * signal because they could carry document text.
 */
export function odfNativeRejection(error: unknown, format: OdfFamily): unknown {
  if (typeof error !== 'string') return error;
  if (error === 'ODF_UNSUPPORTED_ENCRYPTION') return new Error('OOXML_SOURCE_UNSUPPORTED_ENCRYPTION');
  if (error.startsWith('ODF_UNSUPPORTED:')) {
    const reason = error.slice('ODF_UNSUPPORTED:'.length);
    return new Error(`MODEL_SOURCE_UNSUPPORTED:${JSON.stringify({
      format,
      reason: REASON.test(reason) ? reason : 'source-subset',
    })}`);
  }
  return new Error(error);
}
