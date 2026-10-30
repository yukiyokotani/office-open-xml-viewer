import type { OdfFamily } from './odf-source.js';

/**
 * Decode a native reader failure. Only native Result failures enter here; WASM
 * initialization and network errors bypass it. `ODF_UNSUPPORTED:<code>` marks
 * content outside the admitted subset; the model-source v2 signal carries
 * only `reason: 'source-subset'`, so the reader's code is not forwarded.
 * Native messages never reach the signal because they could carry document
 * text.
 */
export function odfNativeRejection(error: unknown, format: OdfFamily): unknown {
  if (typeof error !== 'string') return error;
  if (error === 'ODF_UNSUPPORTED_ENCRYPTION') return new Error('OOXML_SOURCE_UNSUPPORTED_ENCRYPTION');
  if (error.startsWith('ODF_UNSUPPORTED:')) {
    return new Error(`MODEL_SOURCE_UNSUPPORTED:${JSON.stringify({ format, reason: 'source-subset' })}`);
  }
  return new Error(error);
}
