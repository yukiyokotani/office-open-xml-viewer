/**
 * Decode a native reader failure. Only native Result failures enter here; WASM
 * initialization and network errors bypass it. The readers reject only
 * corrupt, encrypted or over-limit input, so the one typed marker is
 * encryption; every other message passes through unchanged (native messages
 * are content-free).
 */
export function odfNativeRejection(error: unknown): unknown {
  if (typeof error !== 'string') return error;
  if (error === 'ODF_UNSUPPORTED_ENCRYPTION') return new Error('OOXML_SOURCE_UNSUPPORTED_ENCRYPTION');
  return new Error(error);
}
