import type { LegacyOfficeConversionOptions } from '../conversion/legacy-office.js';
import {
  MAX_LEGACY_DOC_SOURCE_BYTES,
  validateLegacyDocSourceDescriptor,
} from '../conversion/legacy-doc-source.js';
import {
  MAX_LEGACY_XLS_SOURCE_BYTES,
  validateLegacyXlsSourceDescriptor,
} from '../conversion/legacy-xls-source.js';
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

export type ResolvedDocDocumentInput =
  | Readonly<{ kind: 'ooxml'; bytes: Uint8Array }>
  | Readonly<{
    kind: 'legacy-doc';
    bytes: Uint8Array;
    source: import('../conversion/legacy-doc-source.js').LegacyDocDirectSourceDescriptor;
    signal?: AbortSignal;
  }>;

export async function resolveDocDocumentInput(
  bytes: Uint8Array | ArrayBuffer,
  options?: LegacyOfficeConversionOptions,
  password?: string,
): Promise<ResolvedDocDocumentInput> {
  return resolveNativeInput(
    'doc', 'docx', validateLegacyDocSourceDescriptor,
    MAX_LEGACY_DOC_SOURCE_BYTES, bytes, options, password,
  );
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
  return resolveNativeInput(
    'ppt', 'pptx', validateLegacyPptSourceDescriptor,
    MAX_LEGACY_PPT_SOURCE_BYTES, bytes, options, password,
  );
}

export type ResolvedXlsWorkbookInput =
  | Readonly<{ kind: 'ooxml'; bytes: Uint8Array }>
  | Readonly<{
    kind: 'legacy-xls';
    bytes: Uint8Array;
    source: import('../conversion/legacy-xls-source.js').LegacyXlsDirectSourceDescriptor;
    signal?: AbortSignal;
  }>;

export async function resolveXlsWorkbookInput(
  bytes: Uint8Array | ArrayBuffer,
  options?: LegacyOfficeConversionOptions,
  password?: string,
): Promise<ResolvedXlsWorkbookInput> {
  return resolveNativeInput(
    'xls', 'xlsx', validateLegacyXlsSourceDescriptor,
    MAX_LEGACY_XLS_SOURCE_BYTES, bytes, options, password,
  );
}

async function resolveNativeInput<F extends 'doc' | 'ppt' | 'xls', D>(
  format: F,
  target: 'docx' | 'pptx' | 'xlsx',
  validate: (value: unknown) => D,
  maximum: number,
  bytes: Uint8Array | ArrayBuffer,
  options?: LegacyOfficeConversionOptions,
  password?: string,
): Promise<
  | Readonly<{ kind: 'ooxml'; bytes: Uint8Array }>
  | Readonly<{ kind: `legacy-${F}`; bytes: Uint8Array; source: D; signal?: AbortSignal }>
> {
  const inspected = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const selected = options?.[format];
  if (sniffCfb(inspected) !== 'legacy-binary-format') {
    return { kind: 'ooxml', bytes: await resolveOoxmlContainer(inspected, password) };
  }
  if (!selected || !('source' in selected)) {
    return {
      kind: 'ooxml',
      bytes: await resolveOfficeInputWithOptionalConversion(inspected, target, options, password),
    };
  }
  if ('converter' in selected) {
    throw new TypeError(`legacyConversion.${format} source and converter are mutually exclusive`);
  }
  const source = validate(selected.source);
  const limit = selected.maxInputBytes ?? maximum;
  if (!Number.isSafeInteger(limit) || limit <= 0 || limit > maximum) {
    throw new RangeError(`legacyConversion.${format}.maxInputBytes is invalid`);
  }
  if (sniffLegacyOfficeFormat(inspected) !== format) {
    throw new LegacyOfficeConversionError('unsupported-input', format, target);
  }
  if (inspected.byteLength > limit) {
    throw new LegacyOfficeConversionError('source-too-large', format, target);
  }
  if (selected.signal?.aborted) {
    throw new LegacyOfficeConversionError('aborted', format, target);
  }
  return {
    kind: `legacy-${format}`,
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
