import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { buildCfbWithStreams } from '@silurus/ooxml-core/testing';
import initXlsx, { XlsxArchive } from '../../xlsx/src/wasm/xlsx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { concat, little16, little32 } from './test-fixtures.js';

await initXlsx({ module_or_path: await readFile(new URL('../../xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({
  wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)),
});
const record = (kind: number, data: Uint8Array = new Uint8Array()) => concat(little16(kind), little16(data.length), data);
const bof = (kind: number) => record(0x0809, concat(little16(0x0600), little16(kind), new Uint8Array(12)));

function fixture(flags: number): Uint8Array {
  const bodies = [0, 1].map(index => {
    const number = new Uint8Array(14);
    new DataView(number.buffer).setFloat64(6, index + 1, true);
    return concat(bof(0x0010), record(0x0081, little16(0)), record(0x0203, number), record(0x000a));
  });
  const bound = (offset: number, index: number) => record(0x0085, concat(
    little32(offset), new Uint8Array([index === 0 ? 0 : flags, 0, 1, 0, 65 + index]),
  ));
  const globalsSize = concat(bof(0x0005), bound(0, 0), bound(0, 1), record(0x000a)).length;
  const workbook = concat(bof(0x0005), bound(globalsSize, 0), bound(globalsSize + bodies[0].length, 1),
    record(0x000a), ...bodies);
  return new Uint8Array(buildCfbWithStreams([{ name: 'Workbook', data: workbook }]));
}

const request = (flags: number) => ({
  bytes: fixture(flags), from: 'xls', to: 'xlsx',
  signal: new AbortController().signal, maxOutputBytes: 1024 * 1024,
} as const);

it.each([
  [0x00, undefined], [0x01, 'hidden'], [0x02, 'veryHidden'],
  [0xfc, undefined], [0xfd, 'hidden'], [0xfe, 'veryHidden'],
] as const)('preserves BIFF sheet visibility and ignores unused flag bits: %s', async (flags, visibility) => {
  // MS-XLS 2.4.28 hsState -> ECMA-376 18.2.19 sheet/@state.
  const converted = await converter.convert(request(flags));
  const archive = new XlsxArchive(new Uint8Array(converted.bytes));
  try {
    const { workbook: model } = JSON.parse(new TextDecoder().decode(archive.parse()));
    expect(model.sheets).toHaveLength(2);
    expect(model.sheets.map((sheet: { name: string }) => sheet.name)).toEqual(['A', 'B']);
    expect(model.sheets[0].visibility).toBeUndefined();
    expect(model.sheets[1].visibility).toBe(visibility);
    // Visibility hides a sheet; it must not delete the sheet or its cells.
    for (const index of [1, 2]) {
      const xml = new TextDecoder().decode(archive.extract_image(`xl/worksheets/sheet${index}.xml`));
      expect(xml).toContain(`<v>${index}</v>`);
    }
  } finally { archive.free(); }
});

it.each([0x03, 0xff])('rejects the reserved hsState value instead of revealing the sheet: %s', async flags => {
  await expect(converter.convert(request(flags)).then(() => 'unexpected-success'))
    .rejects.toMatchObject({ reason: 'unsupported-input' });
});
