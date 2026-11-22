/**
 * Calling-realm half of the OpenDocument model sources. A reader plugs into a
 * renderer only through the generic `ModelSource` contract: it claims its own
 * package family from the raw bytes and names a self-contained source module
 * plus its WASM URL.
 */
import type { ModelSource, ModelSourceModuleDescriptor, ModelSourceTarget } from '@silurus/ooxml-core';
import { assertCompanionRuntime, MODEL_SOURCE_MODULE_PROTOCOL } from '@silurus/ooxml/internal/companion';
import { COMPANION_VERSION } from '../../../src/companion-version.js';

export type OdfFamily = 'odt' | 'odp' | 'ods';

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

const TARGETS: Readonly<Record<OdfFamily, ModelSourceTarget>> = { odt: 'docx', odp: 'pptx', ods: 'xlsx' };

/**
 * OpenDocument Part 2 §3.3 media types each family claims. A packaged ODT
 * (Part 3 §2.2.3) must carry the `mimetype` file; templates share the content
 * model of their document type.
 */
const MEDIA_TYPES: Readonly<Record<OdfFamily, readonly string[]>> = {
  odt: ['application/vnd.oasis.opendocument.text', 'application/vnd.oasis.opendocument.text-template'],
  odp: [
    'application/vnd.oasis.opendocument.presentation',
    'application/vnd.oasis.opendocument.presentation-template',
  ],
  ods: ['application/vnd.oasis.opendocument.spreadsheet', 'application/vnd.oasis.opendocument.spreadsheet-template'],
};

const MIMETYPE_NAME = new TextEncoder().encode('mimetype');
/** Implementation bound for the claim: every defined media type is shorter. */
const MAX_MIMETYPE_BYTES = 128;

/**
 * The media type stored in the package's leading `mimetype` file, or
 * `undefined` when the bytes do not start with one. ODF Part 2 §3.3: the
 * `mimetype` file is the first ZIP entry, stored (method 0) and holds the
 * media type in ASCII. Only this first local header is read.
 */
export function leadingOdfMediaType(bytes: Uint8Array): string | undefined {
  if (bytes.byteLength < 30) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint32(0, true) !== 0x04034b50) return undefined;
  const flags = view.getUint16(6, true);
  const method = view.getUint16(8, true);
  const compressed = view.getUint32(18, true);
  const size = view.getUint32(22, true);
  if (!isMimetypeName(bytes, 30, view.getUint16(26, true))) return undefined;
  // Encrypted, stream-sized (data descriptor) or compressed entries are read
  // through the central directory instead.
  if ((flags & 0x0009) !== 0 || method !== 0 || compressed !== size || size === 0 || size > MAX_MIMETYPE_BYTES) {
    return undefined;
  }
  const start = 30 + MIMETYPE_NAME.length + view.getUint16(28, true);
  if (bytes.byteLength < start + size) return undefined;
  let mediaType = '';
  for (let index = start; index < start + size; index++) {
    const byte = bytes[index] as number;
    if (byte < 0x21 || byte > 0x7e) return undefined;
    mediaType += String.fromCharCode(byte);
  }
  return mediaType;
}

/**
 * Whether the package's `mimetype` file holds one of `mediaTypes`. Producers
 * store it first and uncompressed (Part 2 §3.3), which
 * {@link leadingOdfMediaType} reads directly. A package re-zipped by a
 * general archiver may compress it or list it later; LibreOffice opens such
 * packages, and so do the readers. The claim then finds the `mimetype` entry
 * in the central directory and compares its CRC-32 and size with each media
 * type, so nothing is inflated in the calling realm; the reader verifies the
 * content itself.
 */
export function claimsOdfMediaType(bytes: Uint8Array, mediaTypes: ReadonlySet<string>): boolean {
  const leading = leadingOdfMediaType(bytes);
  if (leading !== undefined) return mediaTypes.has(leading);
  const entry = centralMimetypeEntry(bytes);
  if (entry === undefined) return false;
  for (const mediaType of mediaTypes) {
    if (mediaType.length === entry.size && crc32(mediaType) === entry.crc) return true;
  }
  return false;
}

function isMimetypeName(bytes: Uint8Array, at: number, length: number): boolean {
  if (length !== MIMETYPE_NAME.length || bytes.byteLength < at + length) return false;
  for (let index = 0; index < length; index++) {
    if (bytes[at + index] !== MIMETYPE_NAME[index]) return false;
  }
  return true;
}

interface MimetypeEntry { readonly crc: number; readonly size: number }

/**
 * Work bound of one claim: central directory entries visited and ZIP64
 * records kept, over all end-record candidates. Real packages need a few
 * per entry; input that needs more is not claimed (it stays on the OOXML
 * path), so hostile end records cannot stall the calling realm.
 */
const MAX_CLAIM_STEPS = 1 << 18;
const EXHAUSTED: unique symbol = Symbol('claim work exhausted');

/**
 * The CRC-32 and size the central directory records for `mimetype`. Like the
 * reader's ZIP preflight, an end record counts only where its comment ends
 * the input, offsets are relative to the archive start (data may precede
 * it), and ZIP64 records and sizes are resolved; a candidate whose directory
 * does not walk is skipped.
 */
function centralMimetypeEntry(bytes: Uint8Array): MimetypeEntry | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const length = bytes.byteLength;
  const u32 = (at: number) => (at >= 0 && at + 4 <= length ? view.getUint32(at, true) : -1);
  const u64 = (at: number) => (at >= 0 && at + 8 <= length ? Number(view.getBigUint64(at, true)) : -1);
  const ends: number[] = [];
  const locators = new Set<number>();
  for (let end = length - 22; end >= Math.max(0, length - 22 - 0xffff); end--) {
    if (u32(end) !== 0x06054b50 || end + 22 + view.getUint16(end + 20, true) !== length) continue;
    // Single-disk archives only, as the reader's preflight: disk numbers 0.
    if (view.getUint32(end + 4, true) !== 0) continue;
    ends.push(end);
    const locator = end - 20;
    if (u32(locator) === 0x07064b50 && u32(locator + 4) === 0 && u32(locator + 16) === 1) locators.add(locator);
  }
  const budget = { steps: MAX_CLAIM_STEPS };
  // ZIP64 records ending at one of those locators. A record usually sits at
  // the locator's offset (nothing precedes the archive) or right before the
  // locator (no extensible data); only otherwise is the input searched, a
  // search charged to the budget (one step per 16 bytes).
  const records = new Map<number, number[]>();
  const kept = new Set<number>();
  const keep = (at: number) => {
    const size = u64(at + 4);
    if (u32(at) !== 0x06064b50 || size < 44 || !locators.has(at + 12 + size) || u32(at + 16) !== 0 || u32(at + 20) !== 0) return false;
    if (kept.has(at)) return true;
    kept.add(at);
    const list = records.get(at + 12 + size);
    if (list === undefined) records.set(at + 12 + size, [at]); else list.push(at);
    return true;
  };
  let search = false;
  for (const locator of locators) {
    const direct = keep(u64(locator + 8));
    const adjacent = keep(locator - 56);
    search ||= !direct && !adjacent;
  }
  if (search) {
    for (let at = length - 56; at >= 0; at--) {
      if ((at & 15) === 0 && --budget.steps < 0) return undefined;
      if (view.getUint32(at, true) === 0x06064b50 && keep(at) && --budget.steps < 0) return undefined;
    }
  }
  try {
    for (const end of ends) {
      const diskCount = view.getUint16(end + 8, true);
      const count = view.getUint16(end + 10, true);
      const size = view.getUint32(end + 12, true);
      const offset = view.getUint32(end + 16, true);
      const locator = end - 20;
      if (locators.has(locator)) {
        const relative = u64(locator + 8);
        for (const record of records.get(locator) ?? []) {
          // Each count field may be a sentinel independently of the other.
          const resolved = count === 0xffff ? u64(record + 32) : count;
          if ((diskCount === 0xffff ? u64(record + 24) : diskCount) !== resolved) continue;
          const entry = directoryEntry(view, budget, record, record - relative,
            resolved,
            size === 0xffffffff ? u64(record + 40) : size,
            offset === 0xffffffff ? u64(record + 48) : offset);
          if (entry !== false) return entry;
        }
        if (count === 0xffff || size === 0xffffffff || offset === 0xffffffff) continue;
      }
      if (diskCount !== count) continue;
      const entry = directoryEntry(view, budget, end, end - offset - size, count, size, offset);
      if (entry !== false) return entry;
    }
  } catch (error) {
    if (error === EXHAUSTED) return undefined;
    throw error;
  }
  return undefined;
}

/**
 * The `mimetype` entry of the directory at archive-relative `offset` that ends
 * exactly at `directoryEnd`, or `false` when the candidate does not hold one
 * coherent directory.
 */
function directoryEntry(
  view: DataView, budget: { steps: number }, directoryEnd: number, base: number, count: number, size: number, offset: number,
): MimetypeEntry | undefined | false {
  if (base < 0 || count < 0 || size < 0 || offset < 0 || base + offset + size !== directoryEnd) return false;
  return walkCentralDirectory(view, budget, base, base + offset, directoryEnd, count);
}

/** The `mimetype` entry of a directory, or `false` when it does not walk. */
function walkCentralDirectory(
  view: DataView, budget: { steps: number }, base: number, start: number, limit: number, count: number,
): MimetypeEntry | undefined | false {
  const bytes = new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
  let at = start;
  let found: MimetypeEntry | undefined;
  for (let index = 0; index < count; index++) {
    if (--budget.steps < 0) throw EXHAUSTED;
    if (at + 46 > limit || view.getUint32(at, true) !== 0x02014b50) return false;
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const next = at + 46 + nameLength + extraLength + view.getUint16(at + 32, true);
    if (next > limit) return false;
    // Each entry must point at a local header (a ZIP64 offset is resolved by
    // the reader).
    const local = view.getUint32(at + 42, true);
    if (local !== 0xffffffff && (base + local + 30 > view.byteLength || view.getUint32(base + local, true) !== 0x04034b50)) {
      return false;
    }
    if (found === undefined && isMimetypeName(bytes, at + 46, nameLength)) {
      let size = view.getUint32(at + 24, true);
      if (size === 0xffffffff) {
        // ZIP64 extended information (0x0001): the uncompressed size first.
        size = -1;
        const extraEnd = at + 46 + nameLength + extraLength;
        for (let field = at + 46 + nameLength; field + 4 <= extraEnd;) {
          if (--budget.steps < 0) throw EXHAUSTED;
          const fieldSize = view.getUint16(field + 2, true);
          if (field + 4 + fieldSize > extraEnd) break;
          if (view.getUint16(field, true) === 0x0001 && fieldSize >= 8) {
            size = Number(view.getBigUint64(field + 4, true));
            break;
          }
          field += 4 + fieldSize;
        }
      }
      found = { crc: view.getUint32(at + 16, true), size };
    }
    at = next;
  }
  // An optional digital signature record may close the directory.
  if (at + 6 <= limit && view.getUint32(at, true) === 0x05054b50) at += 6 + view.getUint16(at + 4, true);
  return at === limit ? found : false;
}

/** CRC-32 (ZIP, reflected 0xEDB88320) of an ASCII string. */
function crc32(text: string): number {
  let crc = 0xffffffff;
  for (let index = 0; index < text.length; index++) {
    crc ^= text.charCodeAt(index);
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
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
      return claimsOdfMediaType(bytes, mediaTypes);
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
