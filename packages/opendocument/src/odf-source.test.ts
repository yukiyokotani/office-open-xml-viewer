import { crc32, deflateRawSync } from 'node:zlib';
import { describe, expect, it, vi } from 'vitest';
import { validateModelSourceModuleDescriptor } from '@silurus/ooxml-core/internal/model-source';
import { storedZip } from '@silurus/ooxml-core/testing';

// Creating, claiming and describing a source must never load native code.
const loaded = vi.hoisted(() => [] as string[]);
vi.mock('./wasm-odt/opendocument_odt.js', () => { loaded.push('odt glue'); return {}; });
vi.mock('./odt-source-module.ts', () => { loaded.push('odt module'); return {}; });

import { leadingOdfMediaType } from './odf-source.js';
import { createOdt, odt } from './odt.js';
import { buildOdtFixture } from './odt-fixture.js';

const TEXT = 'application/vnd.oasis.opendocument.text';

interface EntryOptions { readonly deflate?: boolean; readonly extra?: boolean }

/** A ZIP as a general archiver writes it: any order, deflate, extra fields. */
function zip(files: ReadonlyArray<readonly [string, string, EntryOptions?]>, comment = ''): Uint8Array<ArrayBuffer> {
  const encoder = new TextEncoder();
  const u16 = (v: number) => [v & 0xff, (v >>> 8) & 0xff];
  const u32 = (v: number) => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
  const local: number[] = [];
  const central: number[] = [];
  for (const [name, content, options = {}] of files) {
    const nameBytes = [...encoder.encode(name)];
    const raw = encoder.encode(content);
    const data = options.deflate ? [...deflateRawSync(raw)] : [...raw];
    const method = options.deflate ? 8 : 0;
    const extra = options.extra ? [...u16(0x5455), ...u16(5), 1, 0, 0, 0, 0] : [];
    const checksum = crc32(raw);
    const offset = local.length;
    local.push(
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(method), ...u16(0), ...u16(0),
      ...u32(checksum), ...u32(data.length), ...u32(raw.length),
      ...u16(nameBytes.length), ...u16(extra.length), ...nameBytes, ...extra, ...data,
    );
    central.push(
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(method), ...u16(0), ...u16(0),
      ...u32(checksum), ...u32(data.length), ...u32(raw.length),
      ...u16(nameBytes.length), ...u16(extra.length), ...u16(0), ...u16(0), ...u16(0),
      ...u32(0), ...u32(offset), ...nameBytes, ...extra,
    );
  }
  const commentBytes = [...encoder.encode(comment)];
  const end = [
    ...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length),
    ...u32(central.length), ...u32(local.length), ...u16(commentBytes.length), ...commentBytes,
  ];
  return Uint8Array.from([...local, ...central, ...end]);
}

/**
 * {@link zip} rewritten as ZIP64: each central size is 0xFFFFFFFF with the
 * real sizes in an extended-information field, and ZIP64 end records whose
 * record carries `extensible` bytes of extensible data.
 */
function zip64(files: ReadonlyArray<readonly [string, string, EntryOptions?]>, extensible = 0): Uint8Array<ArrayBuffer> {
  const plain = zip(files);
  const view = new DataView(plain.buffer);
  const end = plain.length - 22;
  const count = view.getUint16(end + 10, true);
  const directory = view.getUint32(end + 16, true);
  const u16 = (v: number) => [v & 0xff, (v >>> 8) & 0xff];
  const u32 = (v: number) => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
  const u64 = (v: number) => [...u32(v), ...u32(0)];
  const central: number[] = [];
  let at = directory;
  for (let index = 0; index < count; index++) {
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const compressed = view.getUint32(at + 20, true);
    const size = view.getUint32(at + 24, true);
    const header = [...plain.subarray(at, at + 46)];
    header.splice(20, 8, ...u32(0xffffffff), ...u32(0xffffffff));
    header.splice(30, 2, ...u16(extraLength + 20));
    central.push(...header, ...plain.subarray(at + 46, at + 46 + nameLength),
      ...u16(0x0001), ...u16(16), ...u64(size), ...u64(compressed),
      ...plain.subarray(at + 46 + nameLength, at + 46 + nameLength + extraLength));
    at += 46 + nameLength + extraLength;
  }
  const record = directory + central.length;
  const tail = [
    ...u32(0x06064b50), ...u64(44 + extensible), ...u16(45), ...u16(45), ...u32(0), ...u32(0),
    ...u64(count), ...u64(count), ...u64(central.length), ...u64(directory), ...new Array<number>(extensible).fill(0),
    ...u32(0x07064b50), ...u32(0), ...u64(record), ...u32(1),
    ...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(0xffff), ...u16(0xffff),
    ...u32(0xffffffff), ...u32(0xffffffff), ...u16(0),
  ];
  return Uint8Array.from([...plain.subarray(0, directory), ...central, ...tail]);
}

/** A classic archive whose central directory ends with a signature record. */
function withDirectorySignature(plain: Uint8Array<ArrayBuffer>): Uint8Array<ArrayBuffer> {
  const end = plain.length - 22;
  const view = new DataView(plain.buffer);
  const record = [0x50, 0x4b, 0x05, 0x05, 3, 0, 1, 2, 3];
  const tail = plain.slice(end);
  new DataView(tail.buffer).setUint32(12, view.getUint32(end + 12, true) + record.length, true);
  return Uint8Array.from([...plain.subarray(0, end), ...record, ...tail]);
}

it('claims in bounded time when many end records point at ZIP64 locators', () => {
  // ~1,500 locator + end-record candidates after a large body.
  const unit = [0x50, 0x4b, 0x06, 0x07, ...new Array<number>(16).fill(0), 0x50, 0x4b, 0x05, 0x06, ...new Array<number>(18).fill(0)];
  const candidates = 1500;
  const bytes = new Uint8Array(4_000_000 + unit.length * candidates);
  for (let index = 0; index < candidates; index++) {
    const at = 4_000_000 + index * unit.length;
    bytes.set(unit, at);
    // Each end record's comment runs to the end of the input.
    const comment = bytes.length - (at + 20 + 22);
    if (comment <= 0xffff) new DataView(bytes.buffer).setUint16(at + 20 + 20, comment, true);
  }
  const started = performance.now();
  expect(odt.claim(bytes)).toBe(false);
  expect(performance.now() - started).toBeLessThan(500);
});

it('claims in bounded time when many end records reuse one long directory', () => {
  // 65,534 one-letter central entries, then ~2,900 end records whose
  // directories all start at byte 0 and run into the next end record.
  const entries = 65_534;
  const entry = [0x50, 0x4b, 0x01, 0x02, ...new Array<number>(24).fill(0), 1, 0, ...new Array<number>(16).fill(0), 0x61];
  const directory = entries * entry.length;
  const bytes = new Uint8Array(directory + 0xffff);
  for (let index = 0; index < entries; index++) bytes.set(entry, index * entry.length);
  const view = new DataView(bytes.buffer);
  for (let end = directory; end + 22 <= bytes.length; end += 22) {
    view.setUint32(end, 0x06054b50, true);
    view.setUint16(end + 10, 0xfffe, true);
    view.setUint32(end + 12, end, true);
    view.setUint16(end + 20, bytes.length - end - 22, true);
  }
  const started = performance.now();
  expect(odt.claim(bytes)).toBe(false);
  expect(performance.now() - started).toBeLessThan(500);
});

describe('ODT source claim', () => {
  it('claims packaged text documents and templates by their mimetype file', () => {
    expect(odt.claim(buildOdtFixture({ body: '<text:p/>' }))).toBe(true);
    expect(odt.claim(storedZip([['mimetype', `${TEXT}-template`]]))).toBe(true);
    for (const other of ['application/vnd.oasis.opendocument.spreadsheet', 'application/vnd.oasis.opendocument.presentation', `${TEXT}-master`]) {
      expect(odt.claim(storedZip([['mimetype', other]]))).toBe(false);
      expect(odt.claim(zip([['content.xml', '<a/>'], ['mimetype', other, { deflate: true }]]))).toBe(false);
    }
    // Not named mimetype, an OOXML package, or not a ZIP at all.
    expect(odt.claim(storedZip([['mimetypes', TEXT]]))).toBe(false);
    expect(odt.claim(storedZip([['[Content_Types].xml', '<Types/>']]))).toBe(false);
    expect(odt.claim(new Uint8Array([1, 2, 3]))).toBe(false);
    expect(loaded).toEqual([]);
  });

  it('claims re-zipped packages: a compressed, later or extra-field mimetype', () => {
    expect(odt.claim(storedZip([['[Content_Types].xml', '<Types/>'], ['mimetype', TEXT]]))).toBe(true);
    expect(odt.claim(zip([['mimetype', TEXT, { deflate: true }], ['content.xml', '<a/>']]))).toBe(true);
    expect(odt.claim(zip([['content.xml', '<a/>', { deflate: true }], ['mimetype', `${TEXT}-template`, { deflate: true }]]))).toBe(true);
    expect(odt.claim(zip([['mimetype', TEXT, { extra: true }], ['content.xml', '<a/>', { extra: true }]]))).toBe(true);
    // With an archive comment after the central directory.
    expect(odt.claim(zip([['content.xml', '<a/>'], ['mimetype', TEXT, { deflate: true }]], 'comment'))).toBe(true);
    // Data before the archive (offsets stay archive-relative).
    const archive = zip([['content.xml', '<a/>'], ['mimetype', TEXT, { deflate: true }]]);
    expect(odt.claim(Uint8Array.from([...new Uint8Array(64).fill(7), ...archive]))).toBe(true);
    // An end-record signature inside the archive comment is not the record.
    expect(odt.claim(zip([['content.xml', '<a/>'], ['mimetype', TEXT, { deflate: true }]], `x\u0050\u004b\u0005\u0006${'y'.repeat(30)}`))).toBe(true);
    // A comment ending in an end record of another disk (an empty directory).
    const otherDisk = String.fromCharCode(0x50, 0x4b, 5, 6, 1, 0, 1, 0, ...new Array<number>(14).fill(0));
    expect(odt.claim(zip([['content.xml', '<a/>'], ['mimetype', TEXT, { deflate: true }]], `x${otherDisk}`))).toBe(true);
    // A comment ending in an end record whose directory does not walk.
    const fake = String.fromCharCode(0x50, 0x4b, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0);
    expect(odt.claim(zip([['content.xml', '<a/>'], ['mimetype', TEXT, { deflate: true }]], `x${fake}`))).toBe(true);
    // A directory closed by a digital signature record.
    expect(odt.claim(withDirectorySignature(zip([['content.xml', '<a/>'], ['mimetype', TEXT, { deflate: true }]])))).toBe(true);
    // ZIP64: end records and a mimetype size in the extended information.
    expect(odt.claim(zip64([['content.xml', '<a/>'], ['mimetype', TEXT, { deflate: true }]], 8))).toBe(true);
    expect(odt.claim(zip64([['content.xml', '<a/>'], ['mimetype', TEXT, { deflate: true }]]))).toBe(true);
    expect(odt.claim(zip64([['mimetype', `${TEXT}x`, { deflate: true }]]))).toBe(false);
    // A central directory that does not record this media type.
    expect(odt.claim(zip([['mimetype', `${TEXT} `, { deflate: true }]]))).toBe(false);
    expect(loaded).toEqual([]);
  });

  it('reads only a stored, unencrypted leading entry', () => {
    const bytes = storedZip([['mimetype', TEXT]]);
    expect(leadingOdfMediaType(bytes)).toBe(TEXT);
    const compressed = bytes.slice(); compressed[8] = 8;
    expect(leadingOdfMediaType(compressed)).toBeUndefined();
    const encrypted = bytes.slice(); encrypted[6] = 1;
    expect(leadingOdfMediaType(encrypted)).toBeUndefined();
    const truncated = bytes.slice(0, 40);
    expect(leadingOdfMediaType(truncated)).toBeUndefined();
  });

  it('describes a frozen v2 module descriptor with absolute URLs and validates options', () => {
    const source = createOdt({ wasmUrl: 'https://cdn.example.test/odt.wasm', moduleUrl: 'https://cdn.example.test/odt.mjs' });
    const descriptor = validateModelSourceModuleDescriptor(source.beginLoad().module, 'docx');
    expect(descriptor).toMatchObject({ target: 'docx', moduleUrl: 'https://cdn.example.test/odt.mjs', config: { wasmUrl: 'https://cdn.example.test/odt.wasm' } });
    expect(() => createOdt({ wasmUrl: './relative.wasm' })).toThrow(TypeError);
    expect(() => createOdt({ moduleUrl: 'data:text/javascript,0' })).toThrow(TypeError);
    expect(() => createOdt({ maxInputBytes: 1 } as never)).toThrow(TypeError);
    expect(odt.target).toBe('docx');
    expect(Object.isFrozen(odt)).toBe(true);
    expect(loaded).toEqual([]);
  });
});
