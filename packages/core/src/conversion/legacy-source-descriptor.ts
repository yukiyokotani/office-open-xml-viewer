/** Versioned, structured-clone-safe first-party native decoder selection. */
export interface LegacyDirectSourceDescriptor<F extends 'ppt' | 'xls'> {
  readonly protocol: `ooxml-legacy-${F}-source/v1`;
  readonly builtin: F;
  readonly wasmUrl: string;
}

const SUPPORTED_WASM_PROTOCOLS = new Set(['http:', 'https:', 'file:', 'blob:', 'data:']);
const FIELDS = ['protocol', 'builtin', 'wasmUrl'] as const;

/** Validate and detach a structured-clone-safe native source descriptor. */
export function validateLegacySourceDescriptor<F extends 'ppt' | 'xls'>(
  value: unknown,
  format: F,
): LegacyDirectSourceDescriptor<F> {
  const label = `legacy ${format.toUpperCase()} source`;
  const expectedProtocol = `ooxml-legacy-${format}-source/v1` as const;
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${label} descriptor must be an object`);
  }
  const keys = Reflect.ownKeys(value);
  if (keys.length !== FIELDS.length || FIELDS.some(field => !keys.includes(field))) {
    throw new TypeError(`${label} descriptor has unknown or missing fields`);
  }
  const descriptors = Object.getOwnPropertyDescriptors(value);
  if (FIELDS.some(field => !descriptors[field]?.enumerable || !('value' in descriptors[field]))) {
    throw new TypeError(`${label} descriptor fields must be enumerable data properties`);
  }
  const protocol = descriptors.protocol.value as unknown;
  const builtin = descriptors.builtin.value as unknown;
  const wasmUrl = descriptors.wasmUrl.value as unknown;
  if (protocol !== expectedProtocol) {
    throw new TypeError(`unsupported ${label} protocol`);
  }
  if (builtin !== format) {
    throw new TypeError(`unsupported ${label} builtin`);
  }
  if (
    typeof wasmUrl !== 'string' ||
    wasmUrl.length === 0 ||
    wasmUrl.trim() !== wasmUrl
  ) {
    throw new TypeError(`${label} wasmUrl must be a nonempty absolute URL`);
  }
  let url: URL;
  try {
    url = new URL(wasmUrl);
  } catch {
    throw new TypeError(`${label} wasmUrl must be a nonempty absolute URL`);
  }
  if (!SUPPORTED_WASM_PROTOCOLS.has(url.protocol)) {
    throw new TypeError(`${label} wasmUrl protocol is unsupported`);
  }
  return Object.freeze({
    protocol: expectedProtocol,
    builtin: format,
    wasmUrl,
  });
}
