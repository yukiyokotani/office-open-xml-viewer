import { MAX_LEGACY_SOURCE_BYTES } from './legacy-source-limits.js';

/** Read the scalar config a legacy source module receives (fail closed). */
export function readLegacySourceModuleConfig(
  config: unknown,
  label: string,
): Readonly<{ wasmUrl: string; maxInputBytes: number }> {
  if (typeof config !== 'object' || config === null) {
    throw new TypeError(`${label} source module config must be an object`);
  }
  const record = config as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    if (key !== 'wasmUrl' && key !== 'maxInputBytes') {
      throw new TypeError(`${label} source module config has an unknown field`);
    }
  }
  const maxInputBytes = record.maxInputBytes ?? MAX_LEGACY_SOURCE_BYTES;
  if (
    typeof maxInputBytes !== 'number'
    || !Number.isSafeInteger(maxInputBytes)
    || maxInputBytes <= 0
    || maxInputBytes > MAX_LEGACY_SOURCE_BYTES
  ) {
    throw new RangeError(`${label} maxInputBytes is invalid`);
  }
  if (typeof record.wasmUrl !== 'string') {
    throw new TypeError(`${label} wasmUrl must be a nonempty absolute URL`);
  }
  return { wasmUrl: record.wasmUrl, maxInputBytes };
}

/** Independent DOC-only reading capabilities. The common XLS/PPT config
 * remains closed: a Word policy must never cross a sibling source module. */
export interface NativeDocReadingPolicies {
  readonly nativeContourPolicy?: 'strict' | 'relocateForReading';
  readonly nativeWordBreakingPolicy?: 'strict' | 'simplifyForReading';
  readonly nativePictureBulletPolicy?: 'strict' | 'storedSizeForReading';
}
export type ResolvedNativeDocReadingPolicies = Required<NativeDocReadingPolicies>;

export function readNativeDocReadingPolicies(value: unknown): ResolvedNativeDocReadingPolicies {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new TypeError('legacy DOC reading policies must be an object');
  const record = value as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    if (key !== 'nativeContourPolicy' && key !== 'nativeWordBreakingPolicy' && key !== 'nativePictureBulletPolicy')
      throw new TypeError('legacy DOC reading policies have an unknown field');
  }
  const contour = record.nativeContourPolicy === undefined ? 'strict' : record.nativeContourPolicy;
  const wordBreaking = record.nativeWordBreakingPolicy === undefined ? 'strict' : record.nativeWordBreakingPolicy;
  const pictureBullet = record.nativePictureBulletPolicy === undefined ? 'strict' : record.nativePictureBulletPolicy;
  if (contour !== 'strict' && contour !== 'relocateForReading') throw new TypeError('Invalid native DOC contour policy');
  if (wordBreaking !== 'strict' && wordBreaking !== 'simplifyForReading')
    throw new TypeError('legacy DOC nativeWordBreakingPolicy is invalid');
  if (pictureBullet !== 'strict' && pictureBullet !== 'storedSizeForReading')
    throw new TypeError('legacy DOC nativePictureBulletPolicy is invalid');
  return Object.freeze({ nativeContourPolicy: contour, nativeWordBreakingPolicy: wordBreaking,
    nativePictureBulletPolicy: pictureBullet });
}

/** Validate one clone-safe DOC descriptor before any acquisition. Each option
 * reaches its own native constructor slot; selecting one never selects another. */
export function readLegacyDocSourceModuleConfig(config: unknown): Readonly<{
  wasmUrl: string; maxInputBytes: number;
}> & ResolvedNativeDocReadingPolicies {
  if (typeof config !== 'object' || config === null || Array.isArray(config))
    throw new TypeError('legacy DOC source module config must be an object');
  const { nativeContourPolicy, nativeWordBreakingPolicy, nativePictureBulletPolicy, ...common } =
    config as Record<string, unknown>;
  return Object.freeze({
    ...readLegacySourceModuleConfig(common, 'legacy DOC'),
    ...readNativeDocReadingPolicies({ nativeContourPolicy, nativeWordBreakingPolicy, nativePictureBulletPolicy }),
  });
}
