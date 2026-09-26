// MS-XLS 2.4.28 BoundSheet8.hsState -> the XLSX sheet visibility model.
import { expect, it } from 'vitest';
import { concat, little16, little32 } from '../test-fixtures.js';
import { workbookCfb } from '../xls-workbook-fixture.js';
import { testXlsSource } from '../test-sources.js';
import { materializeXlsxWorkbook } from './node-facade.js';

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
  return workbookCfb(concat(bof(0x0005), bound(globalsSize, 0), bound(globalsSize + bodies[0]!.length, 1),
    record(0x000a), ...bodies));
}

const open = (flags: number) => materializeXlsxWorkbook(fixture(flags), { modelSources: [testXlsSource()] });

it.each([
  [0x00, undefined], [0x01, 'hidden'], [0x02, 'veryHidden'],
  [0xfc, undefined], [0xfd, 'hidden'], [0xfe, 'veryHidden'],
] as const)('projects BIFF sheet visibility and ignores unused flag bits: %s', async (flags, visibility) => {
  const { workbookIndex, worksheets } = await open(flags);
  expect(workbookIndex.workbook.sheets.map(sheet => [sheet.name, sheet.visibility])).toEqual([['A', undefined], ['B', visibility]]);
  // Visibility hides a sheet; it must not delete the sheet or its cells.
  expect(worksheets.map(sheet => sheet.rows.flatMap(row => row.cells).map(cell => cell.value)))
    .toEqual([[{ type: 'number', number: 1 }], [{ type: 'number', number: 2 }]]);
});

it.each([0x03, 0xff])('rejects the reserved hsState value instead of revealing the sheet: %s', async flags => {
  await expect(open(flags)).rejects.toThrow(/^UNSUPPORTED:invalid BIFF sheet visibility/);
});
