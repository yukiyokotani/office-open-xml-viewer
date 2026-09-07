/** First-party native decoder asset selected for direct legacy PPT parsing. */
export interface LegacyPptDirectSourceDescriptor {
  readonly protocol: 'ooxml-legacy-ppt-source/v1';
  readonly builtin: 'ppt';
  readonly wasmUrl: string;
}

const SUPPORTED_WASM_PROTOCOLS = new Set(['http:', 'https:', 'file:', 'blob:', 'data:']);
const FIELDS = ['protocol', 'builtin', 'wasmUrl'] as const;

/** Validate and detach a structured-clone-safe native PPT source descriptor. */
export function validateLegacyPptSourceDescriptor(
  value: unknown,
): LegacyPptDirectSourceDescriptor {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('legacy PPT source descriptor must be an object');
  }
  const keys = Reflect.ownKeys(value);
  if (keys.length !== FIELDS.length || FIELDS.some(field => !keys.includes(field))) {
    throw new TypeError('legacy PPT source descriptor has unknown or missing fields');
  }
  const descriptors = Object.getOwnPropertyDescriptors(value);
  if (FIELDS.some(field => !descriptors[field]?.enumerable || !('value' in descriptors[field]))) {
    throw new TypeError('legacy PPT source descriptor fields must be enumerable data properties');
  }
  const protocol = descriptors.protocol.value as unknown;
  const builtin = descriptors.builtin.value as unknown;
  const wasmUrl = descriptors.wasmUrl.value as unknown;
  if (protocol !== 'ooxml-legacy-ppt-source/v1') {
    throw new TypeError('unsupported legacy PPT source protocol');
  }
  if (builtin !== 'ppt') {
    throw new TypeError('unsupported legacy PPT source builtin');
  }
  if (
    typeof wasmUrl !== 'string' ||
    wasmUrl.length === 0 ||
    wasmUrl.trim() !== wasmUrl
  ) {
    throw new TypeError('legacy PPT source wasmUrl must be a nonempty absolute URL');
  }
  let url: URL;
  try {
    url = new URL(wasmUrl);
  } catch {
    throw new TypeError('legacy PPT source wasmUrl must be a nonempty absolute URL');
  }
  if (!SUPPORTED_WASM_PROTOCOLS.has(url.protocol)) {
    throw new TypeError('legacy PPT source wasmUrl protocol is unsupported');
  }
  return Object.freeze({
    protocol,
    builtin,
    wasmUrl,
  });
}
