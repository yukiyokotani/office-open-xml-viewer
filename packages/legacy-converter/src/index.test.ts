import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import initDocx, { docx_to_markdown } from '../../docx/src/wasm/docx_parser.js';
import initXlsx, { xlsx_to_markdown } from '../../xlsx/src/wasm/xlsx_parser.js';
import initPptx, { pptx_to_markdown } from '../../pptx/src/wasm/pptx_parser.js';
import { validateConvertedOoxml } from '@silurus/ooxml-core';
import { createLegacyOfficeWasmConverter } from './index.js';

const converterWasm = readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url));

beforeAll(async () => {
  const [docx, xlsx, pptx] = await Promise.all([
    readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)),
    readFile(new URL('../../xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url)),
    readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)),
  ]);
  await Promise.all([
    initDocx({ module_or_path: docx }),
    initXlsx({ module_or_path: xlsx }),
    initPptx({ module_or_path: pptx }),
  ]);
});

describe('purpose-built legacy Office WASM converter', () => {
  it('converts a Word 97 Unicode main story into parser-readable DOCX', async () => {
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const result = await converter.convert(request('doc', 'docx', buildDocFixture()));
    const bytes = asBytes(result.bytes);

    await validateConvertedOoxml(bytes, 'docx');
    expect(docx_to_markdown(bytes)).toContain('Hello 日本語');
    expect(result).toMatchObject({
      engine: 'silurus-legacy-office',
      engineVersion: '0.1.0',
      warnings: expect.arrayContaining(['legacy-doc:main-story-plain-text-only']),
    });
  });

  it('converts BIFF8 values into parser-readable XLSX', async () => {
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const result = await converter.convert(request('xls', 'xlsx', buildXlsFixture()));
    const bytes = asBytes(result.bytes);

    await validateConvertedOoxml(bytes, 'xlsx');
    const markdown = xlsx_to_markdown(bytes);
    expect(markdown).toContain('表計算');
    expect(markdown).toContain('42.5');
    expect(markdown).toContain('日本語');
  });

  it('converts PowerPoint Unicode text atoms into parser-readable PPTX', async () => {
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const result = await converter.convert(request('ppt', 'pptx', buildPptFixture()));
    const bytes = asBytes(result.bytes);

    await validateConvertedOoxml(bytes, 'pptx');
    expect(pptx_to_markdown(bytes)).toContain('Legacy 日本語 slide');
  });

  it('maps unsupported and bounded-output failures to stable typed reasons', async () => {
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    await expect(converter.convert(request('doc', 'docx', new Uint8Array([1, 2, 3]))))
      .rejects.toMatchObject({ reason: 'unsupported-input' });
    await expect(converter.convert({
      ...request('doc', 'docx', buildDocFixture()),
      maxOutputBytes: 128,
    })).rejects.toMatchObject({ reason: 'output-too-large' });
    await expect(converter.convert({
      ...request('doc', 'docx', buildDocFixture()),
      maxOutputBytes: -1,
    })).rejects.toMatchObject({ reason: 'failed' });
  });
});

function request(
  from: 'doc' | 'xls' | 'ppt',
  to: 'docx' | 'xlsx' | 'pptx',
  bytes: Uint8Array,
) {
  return {
    bytes,
    from,
    to,
    maxOutputBytes: 1024 * 1024,
    signal: new AbortController().signal,
  } as const;
}

function asBytes(bytes: Uint8Array | ArrayBuffer): Uint8Array {
  return bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
}

function buildDocFixture(): Uint8Array {
  const text = 'Hello 日本語\rSecond paragraph';
  const units = Array.from(text, (character) => character.charCodeAt(0));
  const textOffset = 0x400;
  const word = new Uint8Array(textOffset + units.length * 2);
  const view = new DataView(word.buffer);
  view.setUint16(0, 0xa5ec, true);
  view.setUint16(2, 0x00c1, true);
  view.setUint32(0x4c, units.length, true);
  view.setUint32(0x1a2, 0, true);
  view.setUint32(0x1a6, 21, true);
  units.forEach((unit, index) => view.setUint16(textOffset + index * 2, unit, true));
  const table = concat(
    new Uint8Array([0x02]),
    little32(16),
    little32(0),
    little32(units.length),
    little16(0),
    little32(textOffset),
    little16(0),
  );
  return buildCfb([
    ['WordDocument', word],
    ['0Table', table],
  ]);
}

function buildXlsFixture(): Uint8Array {
  const bof = (kind: number) => biffRecord(0x0809, concat(
    little16(0x0600), little16(kind), little16(0), little16(0),
  ));
  const sheetName = utf16le('表計算');
  const boundSheet = concat(
    little32(0), new Uint8Array([0, 0, 3, 1]), sheetName,
  );
  const string = utf16le('日本語');
  const sst = concat(
    little32(1), little32(1), little16(3), new Uint8Array([1]), string,
  );
  const globals = concat(
    bof(0x0005),
    biffRecord(0x0085, boundSheet),
    biffRecord(0x00fc, sst),
    biffRecord(0x000a, new Uint8Array()),
  );
  const number = new Uint8Array(14);
  new DataView(number.buffer).setFloat64(6, 42.5, true);
  const label = concat(little16(1), little16(1), little16(0), little32(0));
  const sheet = concat(
    bof(0x0010),
    biffRecord(0x0203, number),
    biffRecord(0x00fd, label),
    biffRecord(0x000a, new Uint8Array()),
  );
  new DataView(boundSheet.buffer, boundSheet.byteOffset, boundSheet.byteLength)
    .setUint32(0, globals.length, true);
  return buildCfb([['Workbook', concat(
    bof(0x0005),
    biffRecord(0x0085, boundSheet),
    biffRecord(0x00fc, sst),
    biffRecord(0x000a, new Uint8Array()),
    sheet,
  )]]);
}

function buildPptFixture(): Uint8Array {
  const record = (version: number, kind: number, payload: Uint8Array) => concat(
    little16(version), little16(kind), little32(payload.length), payload,
  );
  const documentAtom = concat(little32(5760), little32(4320), new Uint8Array(32));
  const slideReference = concat(little32(2), new Uint8Array(16));
  const document = record(0x000f, 1000, concat(
    record(1, 1001, documentAtom),
    record(15, 4080, record(0, 1011, slideReference)),
  ));
  const slide = record(0x000f, 1006, record(0, 4000, utf16le('Legacy 日本語 slide')));
  const directoryOffset = document.length + slide.length;
  const directory = record(0, 0x1772, concat(little32(0x00200001), little32(0), little32(document.length)));
  const currentEdit = directoryOffset + directory.length;
  const userEdit = record(0, 0x0ff5, concat(new Uint8Array(12), little32(directoryOffset), little32(1), new Uint8Array(8)));
  const currentUserPayload = concat(
    little32(0x14),
    little32(0xe391c05f),
    little32(currentEdit),
    little16(0),
    little16(0x03f4),
    new Uint8Array([3, 0, 0, 0]),
    little32(8),
  );
  const currentUser = record(0, 0x0ff6, currentUserPayload);
  return buildCfb([
    ['PowerPoint Document', concat(document, slide, directory, userEdit)],
    ['Current User', currentUser],
  ]);
}

function biffRecord(kind: number, payload: Uint8Array): Uint8Array {
  return concat(little16(kind), little16(payload.length), payload);
}

function little16(value: number): Uint8Array {
  const bytes = new Uint8Array(2);
  new DataView(bytes.buffer).setUint16(0, value, true);
  return bytes;
}

function little32(value: number): Uint8Array {
  const bytes = new Uint8Array(4);
  new DataView(bytes.buffer).setUint32(0, value, true);
  return bytes;
}

function utf16le(value: string): Uint8Array {
  const bytes = new Uint8Array(value.length * 2);
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < value.length; index++) {
    view.setUint16(index * 2, value.charCodeAt(index), true);
  }
  return bytes;
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function buildCfb(streams: ReadonlyArray<readonly [string, Uint8Array]>): Uint8Array {
  const sectorSize = 512;
  const padded = streams.map(([name, bytes]) => ({
    name,
    bytes,
    declared: Math.max(bytes.length, 4096),
    sectors: Math.ceil(Math.max(bytes.length, 4096) / sectorSize),
  }));
  const dataSectors = padded.reduce((total, entry) => total + entry.sectors, 0);
  const directorySectors = Math.max(1, Math.ceil((streams.length + 1) * 128 / sectorSize));
  let fatSectors = 1;
  while (dataSectors + directorySectors + fatSectors > fatSectors * 128) fatSectors++;
  const output = new Uint8Array(512 + (dataSectors + directorySectors + fatSectors) * sectorSize);
  const view = new DataView(output.buffer);
  output.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  view.setUint16(24, 0x003e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, fatSectors, true);
  view.setUint32(48, dataSectors, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, 0xfffffffe, true);
  view.setUint32(68, 0xfffffffe, true);
  for (let index = 0; index < 109; index++) {
    view.setUint32(76 + index * 4, index < fatSectors
      ? dataSectors + directorySectors + index
      : 0xffffffff, true);
  }

  let sector = 0;
  const starts: Array<{ start: number; size: number }> = [];
  for (const entry of padded) {
    starts.push({ start: sector, size: entry.declared });
    output.set(entry.bytes, 512 + sector * sectorSize);
    sector += entry.sectors;
  }
  const directoryStart = sector;
  const directoryOffset = 512 + directoryStart * sectorSize;
  writeDirectoryEntry(output.subarray(directoryOffset, directoryOffset + 128), 'Root Entry', 5, 0xfffffffe, 0);
  padded.forEach((entry, index) => {
    const offset = directoryOffset + (index + 1) * 128;
    writeDirectoryEntry(output.subarray(offset, offset + 128), entry.name, 2, starts[index].start, starts[index].size);
  });
  sector += directorySectors;
  const fatStart = sector;
  const fat = new Uint32Array(fatSectors * 128).fill(0xffffffff);
  let cursor = 0;
  for (const entry of padded) {
    for (let index = 0; index < entry.sectors; index++) {
      fat[cursor + index] = index + 1 === entry.sectors ? 0xfffffffe : cursor + index + 1;
    }
    cursor += entry.sectors;
  }
  for (let index = 0; index < directorySectors; index++) {
    fat[directoryStart + index] = index + 1 === directorySectors
      ? 0xfffffffe
      : directoryStart + index + 1;
  }
  for (let index = 0; index < fatSectors; index++) {
    fat[fatStart + index] = 0xfffffffd;
    const offset = 512 + (fatStart + index) * sectorSize;
    for (let entry = 0; entry < 128; entry++) {
      view.setUint32(offset + entry * 4, fat[index * 128 + entry], true);
    }
  }
  return output;
}

function writeDirectoryEntry(
  target: Uint8Array,
  name: string,
  objectType: number,
  startSector: number,
  size: number,
): void {
  const view = new DataView(target.buffer, target.byteOffset, target.byteLength);
  for (let index = 0; index < name.length; index++) {
    view.setUint16(index * 2, name.charCodeAt(index), true);
  }
  view.setUint16(name.length * 2, 0, true);
  view.setUint16(64, (name.length + 1) * 2, true);
  target[66] = objectType;
  target[67] = 1;
  view.setUint32(68, 0xffffffff, true);
  view.setUint32(72, 0xffffffff, true);
  view.setUint32(76, 0xffffffff, true);
  view.setUint32(116, startSector, true);
  view.setUint32(120, size, true);
}
