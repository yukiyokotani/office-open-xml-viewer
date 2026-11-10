/**
 * The metric tables of a WOFF or WOFF2 font file, as a minimal sfnt that the
 * OpenType resource parser reads (open-type-metrics.ts). Library resource
 * code, not OOXML behaviour.
 *
 * Only `head`, `hhea`, `OS/2`, `cmap`, `fvar`, `avar` and `MVAR` are kept.
 * WOFF 2.0 transforms only `glyf`, `loca` and `hmtx` (W3C WOFF2 §5), so the
 * kept tables are stored verbatim inside the compressed stream:
 *
 * - WOFF 1.0 (W3C WOFF §5): each table is zlib-compressed on its own and is
 *   inflated with the native `DecompressionStream('deflate')`;
 * - WOFF 2.0: one Brotli stream holds every table, inflated with the native
 *   `DecompressionStream('brotli')` where the engine has it. No Brotli decoder
 *   is bundled: without native support the tables are unavailable
 *   (`'unsupported'`) and callers measure the face instead.
 *
 * One budget per file, {@link MAX_INFLATED_BYTES}, covers everything the
 * decode inflates and allocates: it is checked against the declared sizes
 * before any table buffer is allocated, and against the bytes actually
 * produced while inflating (never trusting a declared size). WOFF 1.0 counts
 * every kept table, inflated or stored; WOFF 2.0 inflates its stream only up
 * to the end of the last kept table. A directory that repeats a tag is
 * malformed. Brotli support is checked before anything is allocated.
 */
import { HARD_MAX_FONT_SOURCE_FACE_BYTES } from '../worker/resource-policy.generated.js';

/** Ceiling on the bytes one file's decode may inflate and allocate in all. */
export const MAX_INFLATED_BYTES = HARD_MAX_FONT_SOURCE_FACE_BYTES * 4;

const KEPT = ['head', 'hhea', 'OS/2', 'cmap', 'fvar', 'avar', 'MVAR'];

const WOFF = 0x774f4646;
const WOFF2 = 0x774f4632;
const SFNT_VERSIONS = new Set([0x00010000, 0x4f54544f /* OTTO */, 0x74727565 /* true */]);

/** WOFF2 §4.1 known table tags, by 6-bit index. */
const WOFF2_TAGS = [
  'cmap', 'head', 'hhea', 'hmtx', 'maxp', 'name', 'OS/2', 'post', 'cvt ', 'fpgm', 'glyf', 'loca', 'prep', 'CFF ',
  'VORG', 'EBDT', 'EBLC', 'gasp', 'hdmx', 'kern', 'LTSH', 'PCLT', 'VDMX', 'vhea', 'vmtx', 'BASE', 'GDEF', 'GPOS',
  'GSUB', 'EBSC', 'JSTF', 'MATH', 'CBDT', 'CBLC', 'COLR', 'CPAL', 'SVG ', 'sbix', 'acnt', 'avar', 'bdat', 'bloc',
  'bsln', 'cvar', 'fdsc', 'feat', 'fmtx', 'fvar', 'gvar', 'hsty', 'just', 'lcar', 'mort', 'morx', 'opbd', 'prop',
  'trak', 'Zapf', 'Silf', 'Glat', 'Gloc', 'Feat', 'Sill',
];

const tagOf = (view: DataView, at: number) => String.fromCharCode(
  view.getUint8(at), view.getUint8(at + 1), view.getUint8(at + 2), view.getUint8(at + 3));

/** Whether this engine inflates Brotli natively. */
export function nativeBrotliAvailable(): boolean {
  try {
    if (typeof DecompressionStream !== 'function') return false;
    new DecompressionStream('brotli' as CompressionFormat);
    return true;
  } catch {
    return false;
  }
}

/** Inflate `input`, keeping only the byte ranges `keep` selects (offsets in
 * the inflated stream), producing at most `limit` bytes; reading stops once
 * `until` bytes were produced. Returns the bytes produced, or null. */
async function inflate(
  format: string,
  input: Uint8Array,
  keep: (offset: number, chunk: Uint8Array) => void,
  limit: number,
  until = Infinity,
): Promise<number | null> {
  let stream: ReadableStream<Uint8Array>;
  try {
    const decompressor = new DecompressionStream(format as CompressionFormat);
    stream = new Blob([input as BlobPart]).stream().pipeThrough(decompressor) as ReadableStream<Uint8Array>;
  } catch {
    return null;
  }
  const reader = stream.getReader();
  let total = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      if (!value) continue;
      if (total + value.byteLength > limit) { await reader.cancel().catch(() => undefined); return null; }
      keep(total, value);
      total += value.byteLength;
      if (total >= until) { await reader.cancel().catch(() => undefined); break; }
    }
  } catch {
    return null;
  }
  return total;
}

/** Copy the part of `chunk` (at stream `offset`) that falls inside a table. */
function collect(table: { start: number; bytes: Uint8Array }, offset: number, chunk: Uint8Array): void {
  const from = Math.max(table.start, offset);
  const to = Math.min(table.start + table.bytes.byteLength, offset + chunk.byteLength);
  if (from < to) table.bytes.set(chunk.subarray(from - offset, to - offset), from - table.start);
}

/** A minimal sfnt holding `tables` (OpenType §5: offset table + records). */
function assembleSfnt(flavor: number, tables: ReadonlyMap<string, Uint8Array>): Uint8Array {
  const tags = [...tables.keys()].sort();
  let size = 12 + tags.length * 16;
  const offsets: number[] = [];
  for (const tag of tags) { offsets.push(size); size += (tables.get(tag)!.byteLength + 3) & ~3; }
  const out = new Uint8Array(size);
  const view = new DataView(out.buffer);
  view.setUint32(0, flavor);
  view.setUint16(4, tags.length);
  tags.forEach((tag, index) => {
    const record = 12 + index * 16;
    for (let char = 0; char < 4; char++) view.setUint8(record + char, tag.charCodeAt(char));
    view.setUint32(record + 8, offsets[index]!);
    view.setUint32(record + 12, tables.get(tag)!.byteLength);
    out.set(tables.get(tag)!, offsets[index]!);
  });
  return out;
}

async function woff1(view: DataView, bytes: Uint8Array): Promise<Uint8Array | null> {
  if (bytes.byteLength < 44) return null;
  const flavor = view.getUint32(4);
  const count = view.getUint16(12);
  if (44 + count * 20 > bytes.byteLength) return null;
  // Directory pass: validate and size every kept table before allocating.
  const kept: { tag: string; stored: Uint8Array; origLength: number }[] = [];
  const seen = new Set<string>();
  let budget = 0;
  for (let index = 0; index < count; index++) {
    const record = 44 + index * 20;
    const tag = tagOf(view, record);
    if (seen.has(tag)) return null;
    seen.add(tag);
    if (!KEPT.includes(tag)) continue;
    const offset = view.getUint32(record + 4);
    const compLength = view.getUint32(record + 8);
    const origLength = view.getUint32(record + 12);
    if (offset + compLength > bytes.byteLength || compLength > origLength) return null;
    budget += origLength;
    if (budget > MAX_INFLATED_BYTES) return null;
    kept.push({ tag, stored: bytes.subarray(offset, offset + compLength), origLength });
  }
  const tables = new Map<string, Uint8Array>();
  for (const { tag, stored, origLength } of kept) {
    if (stored.byteLength === origLength) { tables.set(tag, stored.slice()); continue; }
    // The declared original length sizes the buffer (within the file's
    // budget above); the inflated bytes are counted as they arrive and must
    // match it exactly.
    const table = { start: 0, bytes: new Uint8Array(origLength) };
    const total = await inflate('deflate', stored, (at, chunk) => collect(table, at, chunk), origLength);
    if (total !== origLength) return null;
    tables.set(tag, table.bytes);
  }
  return assembleSfnt(flavor, tables);
}

/** WOFF2 §4.1 UIntBase128. */
function base128(view: DataView, at: { offset: number }, end: number): number | null {
  let value = 0;
  for (let index = 0; index < 5; index++) {
    if (at.offset >= end) return null;
    const byte = view.getUint8(at.offset++);
    if (index === 0 && byte === 0x80) return null;
    if (value & 0xfe000000) return null;
    value = (value << 7) | (byte & 0x7f);
    if ((byte & 0x80) === 0) return value >>> 0;
  }
  return null;
}

async function woff2(view: DataView, bytes: Uint8Array): Promise<Uint8Array | null | 'unsupported'> {
  if (bytes.byteLength < 48) return null;
  const flavor = view.getUint32(4);
  // A collection needs its face directory to pick a face (ttcf flavor).
  if (flavor === 0x74746366) return null;
  const count = view.getUint16(12);
  const compressedLength = view.getUint32(20);
  const cursor = { offset: 48 };
  const wanted: { tag: string; start: number; length: number }[] = [];
  const seen = new Set<string>();
  let streamOffset = 0;
  for (let index = 0; index < count; index++) {
    if (cursor.offset >= bytes.byteLength) return null;
    const flags = view.getUint8(cursor.offset++);
    let tag: string;
    if ((flags & 0x3f) === 0x3f) {
      if (cursor.offset + 4 > bytes.byteLength) return null;
      tag = tagOf(view, cursor.offset);
      cursor.offset += 4;
    } else {
      tag = WOFF2_TAGS[flags & 0x3f] ?? '';
      if (!tag) return null;
    }
    if (seen.has(tag)) return null;
    seen.add(tag);
    const version = flags >> 6;
    const origLength = base128(view, cursor, bytes.byteLength);
    if (origLength === null) return null;
    const transformed = tag === 'glyf' || tag === 'loca' ? version !== 3 : version !== 0;
    const length = transformed ? base128(view, cursor, bytes.byteLength) : origLength;
    if (length === null) return null;
    // Kept tables are never transformed (§5); a transformed one is malformed.
    if (KEPT.includes(tag)) {
      if (transformed) return null;
      wanted.push({ tag, start: streamOffset, length: origLength });
    }
    streamOffset += length;
    if (streamOffset > Number.MAX_SAFE_INTEGER) return null;
  }
  if (cursor.offset + compressedLength > bytes.byteLength) return null;
  if (!nativeBrotliAvailable()) return 'unsupported';
  // The stream is inflated only up to the end of the last kept table (the
  // last chunk may run past it); what is inflated plus the kept buffers
  // stays within the file's budget.
  if (wanted.length === 0) return null;
  const end = wanted.reduce((max, table) => Math.max(max, table.start + table.length), 0);
  const allocated = wanted.reduce((sum, table) => sum + table.length, 0);
  if (end + allocated > MAX_INFLATED_BYTES) return null;
  const tables = wanted.map((table) => ({ tag: table.tag, start: table.start, bytes: new Uint8Array(table.length) }));
  const total = await inflate('brotli', bytes.subarray(cursor.offset, cursor.offset + compressedLength), (at, chunk) => {
    for (const table of tables) collect(table, at, chunk);
  }, MAX_INFLATED_BYTES - allocated, end);
  if (total === null || total < end) return null;
  return assembleSfnt(flavor, new Map(tables.map((table) => [table.tag, table.bytes])));
}

/**
 * The sfnt to parse for `data`: the bytes themselves for TrueType/OpenType
 * (and collections), a minimal sfnt of the metric tables for WOFF/WOFF2,
 * `'unsupported'` for WOFF2 where the engine cannot inflate Brotli, and null
 * for anything malformed or unknown.
 */
export async function metricSfnt(data: ArrayBuffer): Promise<Uint8Array | null | 'unsupported'> {
  const bytes = new Uint8Array(data);
  if (bytes.byteLength < 12) return null;
  const view = new DataView(data);
  const signature = view.getUint32(0);
  if (SFNT_VERSIONS.has(signature) || signature === 0x74746366) return bytes;
  if (signature === WOFF) return woff1(view, bytes);
  if (signature === WOFF2) return woff2(view, bytes);
  return null;
}
