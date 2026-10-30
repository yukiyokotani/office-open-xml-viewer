/**
 * Calling-realm half of the OpenDocument model sources. A reader plugs into a
 * renderer only through the generic `ModelSource` contract: it claims its own
 * package family from the raw bytes and names a self-contained source module
 * plus its WASM URL.
 */
import type { ModelSource, ModelSourceModuleDescriptor, ModelSourceTarget } from '@silurus/ooxml-core';
import { assertCompanionRuntime, MODEL_SOURCE_MODULE_PROTOCOL } from '@silurus/ooxml/internal/companion';
import { COMPANION_VERSION } from '../../../src/companion-version.js';

export type OdfFamily = 'odt';

/** Options accepted by every OpenDocument source factory. */
export interface OpenDocumentSourceOptions {
  /** Absolute URL of the reader's WASM binary. Defaults to the emitted asset. */
  readonly wasmUrl?: string;
  /**
   * Absolute URL of the reader's self-contained ES source module. Defaults to
   * the emitted asset; override it when an asset pipeline relocates files.
   */
  readonly moduleUrl?: string;
}

const URL_PROTOCOLS: Readonly<Record<'wasmUrl' | 'moduleUrl', ReadonlySet<string>>> = {
  wasmUrl: new Set(['http:', 'https:', 'file:', 'blob:', 'data:']),
  moduleUrl: new Set(['http:', 'https:', 'file:', 'blob:']),
};

const TARGETS: Readonly<Record<OdfFamily, ModelSourceTarget>> = { odt: 'docx' };

/**
 * OpenDocument Part 2 §3.3 media types each family claims. A packaged ODT
 * (Part 3 §2.2.3) must carry the `mimetype` file; templates share the content
 * model of their document type.
 */
const MEDIA_TYPES: Readonly<Record<OdfFamily, readonly string[]>> = {
  odt: ['application/vnd.oasis.opendocument.text', 'application/vnd.oasis.opendocument.text-template'],
};

const MIMETYPE_NAME = new TextEncoder().encode('mimetype');
/** Implementation bound for the claim: every defined media type is shorter. */
const MAX_MIMETYPE_BYTES = 128;

/**
 * The media type stored in the package's leading `mimetype` file, or
 * `undefined` when the bytes do not start with one. ODF Part 2 §3.3: the
 * `mimetype` file is the first ZIP entry, stored (method 0) and holds the
 * media type in ASCII. Only this first local header is read; the owning realm
 * validates the archive, the manifest and the absence of an extra field.
 */
export function leadingOdfMediaType(bytes: Uint8Array): string | undefined {
  if (bytes.byteLength < 30) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint32(0, true) !== 0x04034b50) return undefined;
  const flags = view.getUint16(6, true);
  const method = view.getUint16(8, true);
  const compressed = view.getUint32(18, true);
  const size = view.getUint32(22, true);
  const nameLength = view.getUint16(26, true);
  const extraLength = view.getUint16(28, true);
  // Encrypted, stream-sized (data descriptor) or compressed entries are not a
  // conforming `mimetype` file.
  if ((flags & 0x0009) !== 0 || method !== 0 || compressed !== size || size === 0 || size > MAX_MIMETYPE_BYTES) {
    return undefined;
  }
  if (nameLength !== MIMETYPE_NAME.length) return undefined;
  const start = 30 + nameLength + extraLength;
  if (bytes.byteLength < start + size) return undefined;
  for (let index = 0; index < nameLength; index++) {
    if (bytes[30 + index] !== MIMETYPE_NAME[index]) return undefined;
  }
  let mediaType = '';
  for (let index = start; index < start + size; index++) {
    const byte = bytes[index] as number;
    if (byte < 0x21 || byte > 0x7e) return undefined;
    mediaType += String.fromCharCode(byte);
  }
  return mediaType;
}

export function createOdfSource<T extends ModelSourceTarget>(
  family: OdfFamily,
  options: OpenDocumentSourceOptions,
  defaults: Readonly<{ wasmUrl: string; moduleUrl: string }>,
): ModelSource<T> {
  // The installed viewer owns protocol and transport. Pin our compiled
  // expectation too: copying its current protocol would hide a stale add-on.
  assertCompanionRuntime(COMPANION_VERSION, 'ooxml-model-source-module/v2');
  const label = `OpenDocument ${family.toUpperCase()} source`;
  if (typeof options !== 'object' || options === null || Array.isArray(options)) {
    throw new TypeError(`${label} options must be an object`);
  }
  for (const key of Object.keys(options)) {
    if (key !== 'wasmUrl' && key !== 'moduleUrl') throw new TypeError(`${label} option ${key} is unknown`);
  }
  const wasmUrl = absoluteUrl(options.wasmUrl ?? defaults.wasmUrl, 'wasmUrl', label);
  const moduleUrl = absoluteUrl(options.moduleUrl ?? defaults.moduleUrl, 'moduleUrl', label);
  const target = TARGETS[family] as T;
  const mediaTypes = new Set(MEDIA_TYPES[family]);
  const module: ModelSourceModuleDescriptor = Object.freeze({
    protocol: MODEL_SOURCE_MODULE_PROTOCOL,
    target,
    moduleUrl,
    config: Object.freeze({ wasmUrl }),
  });
  return Object.freeze({
    target,
    claim(bytes: Uint8Array): boolean {
      const mediaType = leadingOdfMediaType(bytes);
      return mediaType !== undefined && mediaTypes.has(mediaType);
    },
    beginLoad() {
      // The calling realm retains no transfer or native handle; the opened
      // package belongs to the source module and is closed by its owner.
      return { module, release() {} };
    },
  });
}

/**
 * A source with the bundled defaults that can be exported as a constant. The
 * real source (companion check and asset URL resolution) is created on the
 * first `claim()` or `beginLoad()`, never at import time.
 */
export function deferredOdfSource<T extends ModelSourceTarget>(
  target: T,
  create: () => ModelSource<T>,
): ModelSource<T> {
  let source: ModelSource<T> | undefined;
  const resolve = (): ModelSource<T> => (source ??= create());
  return Object.freeze({
    target,
    claim: (bytes: Uint8Array): boolean => resolve().claim(bytes),
    beginLoad: () => resolve().beginLoad(),
  });
}

function absoluteUrl(value: unknown, field: 'wasmUrl' | 'moduleUrl', label: string): string {
  if (typeof value !== 'string' || value.length === 0 || value.trim() !== value) {
    throw new TypeError(`${label} ${field} must be a nonempty absolute URL`);
  }
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new TypeError(`${label} ${field} must be a nonempty absolute URL`);
  }
  if (!URL_PROTOCOLS[field].has(url.protocol)) {
    throw new TypeError(`${label} ${field} protocol is unsupported`);
  }
  return value;
}
