// BIFF8 Window1/Window2 (MS-XLS 2.4.345/2.4.346) through the direct XLS reader
// into the XLSX worksheet view model. The full flag matrix and the window
// association limits are unit-tested in xls/views.rs.
import { expect, it } from 'vitest';
import { concat, little16, little32 } from '../test-fixtures.js';
import { workbookCfb } from '../xls-workbook-fixture.js';
import { testXlsSource } from '../test-sources.js';
import { materializeXlsxWorkbook } from './node-facade.js';

const record = (kind: number, data: Uint8Array = new Uint8Array()) => concat(little16(kind), little16(data.length), data);
const bof = (kind: number) => record(0x0809, concat(little16(0x0600), little16(kind), new Uint8Array(12)));
const window2 = (flags: number) => record(0x023e, concat(little16(flags), new Uint8Array(16)));
const GRIDLINES = 0x02, HEADERS = 0x04, ZEROS = 0x10, RTL = 0x40;

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
  return workbookCfb(concat(bof(0x0005), ...windows, bound(size), record(0x000a), body));
}

async function view(bytes: Uint8Array) {
  const [worksheet] = (await materializeXlsxWorkbook(bytes, { modelSources: [testXlsSource()] })).worksheets;
  expect(worksheet!.rows.flatMap(row => row.cells).map(cell => cell.value)).toEqual([{ type: 'number', number: 0 }]);
  return { gridlines: worksheet!.showGridlines, zeros: worksheet!.showZeros, rtl: worksheet!.rightToLeft };
}

// The XLSX model carries no row/column-header flag, so HEADERS must not leak
// into any projected display bit.
it.each([
  [GRIDLINES | HEADERS | ZEROS | RTL, { gridlines: true, zeros: true, rtl: true }],
  [0, { gridlines: false, zeros: false, rtl: false }],
  [GRIDLINES, { gridlines: true, zeros: false, rtl: false }],
  [HEADERS | ZEROS | RTL, { gridlines: false, zeros: true, rtl: true }],
])('projects Window2 display flags %s without dropping cells', async (flags, expected) => {
  expect(await view(fixture([flags]))).toEqual(expected);
});

it('projects the last worksheet window of several workbook windows', async () => {
  expect(await view(fixture([GRIDLINES | HEADERS | ZEROS, RTL]))).toEqual({ gridlines: false, zeros: false, rtl: true });
});

it('does not import Window2 from embedded charts or saved custom views', async () => {
  const extra = concat(bof(0x0020), record(0x023e, new Uint8Array(10)), record(0x000a),
    record(0x01aa), window2(0), record(0x01ab));
  expect(await view(fixture([GRIDLINES | HEADERS | ZEROS], { extra }))).toEqual({ gridlines: true, zeros: true, rtl: false });
});

it('keeps the XLSX defaults when neither workbook nor sheet windows are supplied', async () => {
  expect(await view(fixture([]))).toEqual({ gridlines: true, zeros: true, rtl: false });
});

it.each([
  fixture([0], { count: 0 }), fixture([], { count: 1 }), fixture([0], { count: 2 }),
  fixture([0], { windowLength: 17 }),
  fixture([], { count: 1, extra: record(0x023e, new Uint8Array(10)) }),
  fixture(Array.from({ length: 1025 }, () => 0)),
])('rejects malformed or unbound views rather than inventing associations (%#)', async bytes => {
  await expect(view(bytes)).rejects.toThrow(/^UNSUPPORTED:.*BIFF (workbook|worksheet) windows/);
});
