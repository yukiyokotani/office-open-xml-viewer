import type { LegacyOfficeConversionOptions } from '../conversion/legacy-office.js';
import { resolveOoxmlContainer } from '../errors/cfb-guard.js';
import type { OoxmlFormat } from '../errors/ooxml-error.js';

/**
 * Preserve the existing OOXML/decryption promise when conversion is omitted,
 * and load the heavier converter boundary only after an explicit opt-in.
 */
export function resolveOfficeInputWithOptionalConversion(
  bytes: Uint8Array | ArrayBuffer,
  target: OoxmlFormat,
  options?: LegacyOfficeConversionOptions,
  password?: string,
): Promise<Uint8Array> {
  if (options === undefined) return resolveOoxmlContainer(bytes, password);
  return import('../conversion/legacy-office.js')
    .then(({ normalizeOfficeInput }) => normalizeOfficeInput(bytes, target, options, password))
    .then((result) => result.bytes);
}

/** Bind one owner/session cancellation signal to an optional converter request. */
export function bindLegacyOfficeConversionSignal(
  options: LegacyOfficeConversionOptions | undefined,
  lifecycleSignal: AbortSignal | undefined,
): Readonly<{
  options?: LegacyOfficeConversionOptions;
  cleanup: () => void;
}> {
  if (options === undefined) return { cleanup: () => {} };
  const combined = combineAbortSignals(options.signal, lifecycleSignal);
  return {
    options: {
      ...options,
      ...(combined.signal === undefined ? {} : { signal: combined.signal }),
    },
    cleanup: combined.cleanup,
  };
}

function combineAbortSignals(
  first: AbortSignal | undefined,
  second: AbortSignal | undefined,
): Readonly<{ signal?: AbortSignal; cleanup: () => void }> {
  if (first === undefined || first === second) {
    return { signal: second, cleanup: () => {} };
  }
  if (second === undefined) return { signal: first, cleanup: () => {} };

  const controller = new AbortController();
  const abort = (): void => controller.abort();
  if (first.aborted || second.aborted) {
    abort();
    return { signal: controller.signal, cleanup: () => {} };
  }
  first.addEventListener('abort', abort, { once: true });
  second.addEventListener('abort', abort, { once: true });
  return {
    signal: controller.signal,
    cleanup: () => {
      first.removeEventListener('abort', abort);
      second.removeEventListener('abort', abort);
    },
  };
}
