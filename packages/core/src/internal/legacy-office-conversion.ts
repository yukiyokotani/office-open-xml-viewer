import type { LegacyOfficeConversionOptions } from '../conversion/legacy-office.js';
import { LegacyOfficeConversionError } from '../conversion/legacy-office-error.js';
import {
  MAX_LEGACY_PPT_SOURCE_BYTES,
  validateLegacyPptSourceDescriptor,
} from '../conversion/legacy-ppt-source.js';
import { sniffCfb, sniffLegacyOfficeFormat } from '../errors/cfb-sniff.js';
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
  if (options?.[legacyFormatForTarget(target)] === undefined) {
    return resolveOoxmlContainer(bytes, password);
  }
  return import('../conversion/legacy-office.js')
    .then(({ normalizeOfficeInput }) => normalizeOfficeInput(bytes, target, options, password))
    .then((result) => result.bytes);
}

export type ResolvedPptPresentationInput =
  | Readonly<{ kind: 'ooxml'; bytes: Uint8Array }>
  | Readonly<{
    kind: 'legacy-ppt';
    bytes: Uint8Array;
    source: import('../conversion/legacy-ppt-source.js').LegacyPptDirectSourceDescriptor;
    signal?: AbortSignal;
  }>;

export async function resolvePptPresentationInput(
  bytes: Uint8Array | ArrayBuffer,
  options?: LegacyOfficeConversionOptions,
  password?: string,
): Promise<ResolvedPptPresentationInput> {
  const inspected = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const selected = options?.ppt;
  if (sniffCfb(inspected) !== 'legacy-binary-format') {
    return { kind: 'ooxml', bytes: await resolveOoxmlContainer(inspected, password) };
  }
  if (!selected || !('source' in selected)) {
    return {
      kind: 'ooxml',
      bytes: await resolveOfficeInputWithOptionalConversion(inspected, 'pptx', options, password),
    };
  }
  if ('converter' in selected) {
    throw new TypeError('legacyConversion.ppt source and converter are mutually exclusive');
  }
  const source = validateLegacyPptSourceDescriptor(selected.source);
  const limit = selected.maxInputBytes ?? MAX_LEGACY_PPT_SOURCE_BYTES;
  if (!Number.isSafeInteger(limit) || limit <= 0 || limit > MAX_LEGACY_PPT_SOURCE_BYTES) {
    throw new RangeError('legacyConversion.ppt.maxInputBytes is invalid');
  }
  if (sniffLegacyOfficeFormat(inspected) !== 'ppt') {
    throw new LegacyOfficeConversionError('unsupported-input', 'ppt', 'pptx');
  }
  if (inspected.byteLength > limit) {
    throw new LegacyOfficeConversionError('source-too-large', 'ppt', 'pptx');
  }
  if (selected.signal?.aborted) {
    throw new LegacyOfficeConversionError('aborted', 'ppt', 'pptx');
  }
  return {
    kind: 'legacy-ppt',
    bytes: inspected.byteOffset === 0 && inspected.byteLength === inspected.buffer.byteLength
      ? inspected
      : inspected.slice(),
    source,
    ...(selected.signal ? { signal: selected.signal } : {}),
  };
}

/** Bind one owner/session cancellation signal to an optional converter request. */
export function bindLegacyOfficeConversionSignal(
  options: LegacyOfficeConversionOptions | undefined,
  target: OoxmlFormat,
  lifecycleSignal: AbortSignal | undefined,
): Readonly<{
  options?: LegacyOfficeConversionOptions;
  cleanup: () => void;
}> {
  if (options === undefined) return { cleanup: () => {} };
  const format = legacyFormatForTarget(target);
  const selected = options[format];
  if (selected === undefined) return { options, cleanup: () => {} };
  const combined = combineAbortSignals(selected.signal, lifecycleSignal);
  return {
    options: {
      ...options,
      [format]: {
        ...selected,
        ...(combined.signal === undefined ? {} : { signal: combined.signal }),
      },
    },
    cleanup: combined.cleanup,
  };
}

function legacyFormatForTarget(target: OoxmlFormat): 'doc' | 'xls' | 'ppt' {
  switch (target) {
    case 'docx': return 'doc';
    case 'xlsx': return 'xls';
    case 'pptx': return 'ppt';
  }
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
