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
