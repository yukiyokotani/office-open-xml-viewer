// Binary Word table shading and positioning (MS-DOC 2.9.253 Shd,
// 2.6.5 sprmTDefTableShd/sprmTShdTable, sprmTPc/sprmTDxaAbs...) as TTP-mark
// row properties through the direct DOC reader.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { materializeDocxDocument } from './node-facade.js';

// Two cells and a row mark; the TTP mark owns the row properties (a two-cell
// sprmTInsert of 1000 twips each plus the tested sprm). Body cell marks carry
// only sprmPFInTable.
const cell = new Uint8Array([0, 0, 0x16, 0x24, 1]);
const tableFixture = (code: number, operand: Uint8Array) => buildDocFixture({ text: 'A\x07B\x07\x07\r', paragraphMarks: [
  { end: 2, properties: cell }, { end: 4, properties: cell },
  { end: 5, properties: concat(cell, new Uint8Array([0x17, 0x24, 1, 0x21, 0x76, 0, 2, 0xe8, 3]), little16(code), operand) },
] });
const modern = (pattern = 0) => concat(new Uint8Array([0, 0, 0, 255, 0x12, 0x34, 0x56, 0]), little16(pattern));
const load = (bytes: Uint8Array) => materializeDocxDocument(bytes, { modelSources: [testDocSource()] });
interface Table { type: string; tblpPr?: unknown; overlap?: string; colWidths: number[]; rows: { cells: { background: string | null }[] }[] }
const tableOf = async (bytes: Uint8Array) => (await load(bytes)).body.find(block => (block as { type: string }).type === 'table') as unknown as Table;

it('retains clear cell shading as the first cell background only', async () => {
  const table = await tableOf(tableFixture(0xd612, concat(new Uint8Array([10]), modern(0))));
  expect(table.colWidths).toEqual([50, 50]);
  expect(table.rows.map(row => row.cells.map(c => c.background))).toEqual([['123456', null]]);
});

it.each([
  [1, 'cannot resolve automatic solid cell shading'],
  [8, 'cannot retain patterned cell shading'],
  [0x23, 'encountered unsupported formatting'],
])('direct rejects cell shading pattern %i', async (pattern, reason) => {
  // The converter wrote w:shd patterns (and warned and dropped the unmappable
  // 0x23); the DOCX cell model has only a resolved solid background, and a
  // solid pattern takes the automatic foreground here.
  await expect(load(tableFixture(0xd612, concat(new Uint8Array([10]), modern(pattern as number))))).rejects.toThrow(new Error(`UNSUPPORTED:direct DOC model ${reason}`));
});

it('direct rejects table-wide shading instead of synthesizing cell overrides', async () => {
  await expect(load(tableFixture(0xd660, concat(new Uint8Array([10]), modern())))).rejects.toThrow(new Error('UNSUPPORTED:direct DOC model cannot retain table-level shading'));
});

it.each([
  [0xd612, new Uint8Array([1, 0]), 'invalid Word cell shading array'],
  [0xd660, new Uint8Array([9, ...new Uint8Array(9)]), 'invalid Word table shading length'],
  [0xd62d, concat(new Uint8Array([12, 0, 3]), modern()), 'Word cell range outside row'],
  [0xd609, new Uint8Array([2, 31, 0]), 'invalid Word shading palette index'],
])('rejects malformed shading operands, code=%i', async (code, operand, reason) => {
  await expect(load(tableFixture(code as number, operand as Uint8Array))).rejects.toThrow(new Error(`UNSUPPORTED:${reason as string}`));
});

it('retains the table position and no-overlap', async () => {
  const table = await tableOf(tableFixture(0x360d, concat(new Uint8Array([0x20]),
    little16(0x940e), little16(721), little16(0x940f), little16(1),
    little16(0x941e), little16(187), little16(0x3465), new Uint8Array([1]))));
  expect(table.tblpPr).toMatchObject({ horzAnchor: 'text', vertAnchor: 'text', tblpX: 36, tblpY: 0, rightFromText: 9.35 });
  expect(table.overlap).toBe('never');
});

// The full PositionCodeOperand byte domain is Rust-covered
// (doc::table::position::anchors_and_ignored_padding_cover_the_complete_byte_domain).
it('keeps an explicitly non-positioned anchor code inline', async () => {
  expect((await tableOf(tableFixture(0x360d, new Uint8Array([0x30])))).tblpPr).toBeUndefined();
});

it.each([
  [0x940e, 'position'],
  [0x9410, 'wrap distance'],
])('rejects out-of-range table position sprm %i', async (code, kind) => {
  await expect(load(tableFixture(code as number, little16(32767)))).rejects.toThrow(new Error(`UNSUPPORTED:invalid Word table ${kind}`));
});
