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
const window2 = (flags: number) => record(0x023e, concat(little16(flags), new Uint8Array(16)));

function fixture(flags: number[], options: { count?: number; extra?: Uint8Array; windowLength?: number } = {}): Uint8Array {
  const number = new Uint8Array(14); // A1 numeric zero must survive even when hidden.
  const body = concat(bof(0x0010), record(0x0203, number), ...flags.map(window2),
    options.extra ?? new Uint8Array(), record(0x000a));
  const windows = Array.from({ length: options.count ?? flags.length }, () =>
    record(0x003d, options.windowLength === undefined
      ? concat(little16(0), little16(0), little16(2000), little16(1000),
        little16(0x38), little16(0), little16(0), little16(1), little16(600))
      : new Uint8Array(options.windowLength)));
  const bound = (offset: number) => record(0x0085, concat(little32(offset), new Uint8Array([0, 0, 1, 0, 65])));
  const size = concat(bof(0x0005), ...windows, bound(0), record(0x000a)).length;
  return new Uint8Array(buildCfbWithStreams([{ name: 'Workbook', data:
    concat(bof(0x0005), ...windows, bound(size), record(0x000a), body),
  }]));
}

async function convert(bytes: Uint8Array) {
  const output = await converter.convert({ bytes, from: 'xls', to: 'xlsx',
    signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 });
  const archive = new XlsxArchive(new Uint8Array(output.bytes));
  try {
    const decode = (part: string) => new TextDecoder().decode(archive.extract_image(part));
    archive.parse();
    archive.open_sheet_cursor(0, 'A');
    let model;
    const rows: unknown[] = [];
    for (let pulls = 0; pulls < 10; pulls++) {
      const product = JSON.parse(new TextDecoder().decode(archive.pull_sheet_cursor(100)));
      if (archive.sheet_cursor_pull_finished()) {
        expect(product.kind).toBe('finished');
        model = product.worksheet;
        archive.acknowledge_sheet_cursor_terminal();
        break;
      }
      rows.push(...product.rows);
    }
    expect(model).toBeDefined();
    archive.close_sheet_cursor();
    model.rows = rows;
    return { workbook: decode('xl/workbook.xml'), sheet: decode('xl/worksheets/sheet1.xml'),
      model };
  } finally { archive.free(); }
}

it.each(Array.from({ length: 16 }, (_, n) => n))('preserves every display-flag combination through XLSX parsing: %s', async n => {
  const flags = (n & 1 ? 2 : 0) | (n & 2 ? 4 : 0) | (n & 4 ? 16 : 0) | (n & 8 ? 64 : 0);
  const { workbook, sheet, model } = await convert(fixture([flags]));
  expect(workbook).toContain('<bookViews><workbookView/></bookViews><sheets>');
  expect(sheet).toContain(`showGridLines="${Number(Boolean(n & 1))}"`);
  expect(sheet).toContain(`showRowColHeaders="${Number(Boolean(n & 2))}"`);
  expect(sheet).toContain(`showZeros="${Number(Boolean(n & 4))}"`);
  expect(sheet).toContain(`rightToLeft="${Number(Boolean(n & 8))}"`);
  expect(model.showGridlines).toBe(Boolean(n & 1));
  expect(model.showZeros).toBe(Boolean(n & 4));
  expect(model.rightToLeft).toBe(Boolean(n & 8));
  expect(sheet).toContain('<c r="A1" s="0"><v>0</v></c>');
  expect(sheet.indexOf('</sheetViews>')).toBeLessThan(sheet.indexOf('<sheetData>'));
});

it('retains ordinal workbook-window associations instead of collapsing multiple views', async () => {
  const { workbook, sheet } = await convert(fixture([0x16, 0x40]));
  expect(workbook.match(/<workbookView\/>/g)).toHaveLength(2);
  expect(sheet).toContain('<sheetView workbookViewId="0" showGridLines="1" showRowColHeaders="1" showZeros="1" rightToLeft="0"/>');
  expect(sheet).toContain('<sheetView workbookViewId="1" showGridLines="0" showRowColHeaders="0" showZeros="0" rightToLeft="1"/>');
});

it('does not import Window2 from embedded charts or saved custom views', async () => {
  const extra = concat(bof(0x0020), record(0x023e, new Uint8Array(10)), record(0x000a),
    record(0x01aa), window2(0), record(0x01ab));
  const { sheet } = await convert(fixture([0x16], { extra }));
  expect(sheet.match(/<sheetView /g)).toHaveLength(1);
  expect(sheet).toContain('showGridLines="1"');
});

it('keeps the previous default when neither workbook nor sheet windows are supplied', async () => {
  const { workbook, sheet, model } = await convert(fixture([]));
  expect(workbook).not.toContain('<bookViews>');
  expect(sheet).not.toContain('<sheetViews>');
  expect(model.showGridlines).toBe(true);
  expect(model.showZeros).toBe(true);
  expect(model.rightToLeft).toBe(false);
});

it.each([
  fixture([0], { count: 0 }), fixture([], { count: 1 }), fixture([0], { count: 2 }),
  fixture([0], { windowLength: 17 }),
  fixture([], { count: 1, extra: record(0x023e, new Uint8Array(10)) }),
  fixture(Array.from({ length: 1025 }, () => 0)),
])('rejects malformed or unbound views rather than inventing associations', async bytes => {
  await expect(convert(bytes)).rejects.toMatchObject({ reason: 'unsupported-input' });
});
