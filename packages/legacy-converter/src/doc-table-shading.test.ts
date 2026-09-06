import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initDocx, { DocxArchive } from '../../docx/src/wasm/docx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildDocFixture, concat, little16 } from './test-fixtures.js';

await initDocx({ module_or_path: await readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });

// Add a physical PAPX FKP to the existing synthetic fixture's reserved area.
// The TTP mark owns the row properties; body cell marks do not carry shading.
function tableFixture(code: number, operand: Uint8Array): Uint8Array {
  const bytes = buildDocFixture({ text: 'A\x07B\x07\x07\r' });
  const word = bytes.subarray(512, 512 + 4096);
  const table = bytes.subarray(512 + 4096, 512 + 8192);
  const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
  view.setUint32(0x102, 128, true);
  view.setUint32(0x106, 12, true);
  const plc = new DataView(table.buffer, table.byteOffset + 128, 12);
  plc.setUint32(0, 1024, true); plc.setUint32(4, 1036, true); plc.setUint32(8, 3, true);
  // Page 1 overlaps the extended FIB; page 2 contains text. Use page 3.
  const page = word.subarray(1536, 2048);
  const pv = new DataView(page.buffer, page.byteOffset, page.byteLength);
  [1024, 1028, 1032, 1034, 1036].forEach((fc, i) => pv.setUint32(i * 4, fc, true));
  const cell = new Uint8Array([0, 0, 0x16, 0x24, 1]);
  const row = concat(cell, new Uint8Array([0x17, 0x24, 1, 0x21, 0x76, 0, 2, 0xe8, 3]), little16(code), operand);
  let offset = 128;
  [cell, cell, row, new Uint8Array()].forEach((props, i) => {
    if (!props.length) return;
    page[20 + i * 13] = offset / 2;
    const header = props.length % 2 === 0 ? 2 : 1;
    page[offset + header - 1] = Math.ceil(props.length / 2);
    page.set(props, offset + header);
    offset = (offset + header + props.length + 1) & ~1;
  });
  page[511] = 4;
  return bytes;
}
const modern = (pattern = 0) => concat(new Uint8Array([0, 0, 0, 255, 0x12, 0x34, 0x56, 0]), little16(pattern));
const convert = (bytes: Uint8Array) => converter.convert({ bytes, from: 'doc', to: 'docx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 });

it.each([[0, 'clear'], [1, 'solid'], [8, 'pct50'], [0x15, 'thinVertStripe'], [0x25, 'pct12']])(
  'retains binary cell shading pattern %i as %s in a parseable package', async (pattern, name) => {
    const result = await convert(tableFixture(0xd612, concat(new Uint8Array([10]), modern(pattern as number))));
    const archive = new DocxArchive(new Uint8Array(result.bytes));
    try {
      const xml = new TextDecoder().decode(archive.extract_image('word/document.xml'));
      expect(xml).toContain(`<w:shd w:val="${name}" w:color="auto" w:fill="123456"/>`);
      expect(xml.match(/<w:shd /g)).toHaveLength(1);
      const model = JSON.parse(new TextDecoder().decode(archive.parse())) as {
        body: { type: string; rows?: { cells: { background: string | null }[] }[] }[];
      };
      if (pattern === 0) {
        const table = model.body.find(block => block.type === 'table');
        expect(table?.rows?.[0]?.cells.map(cell => cell.background)).toEqual(['123456', null]);
      }
    } finally { archive.free(); }
  },
);

it('retains table-wide shading without synthesizing direct cell overrides', async () => {
  const result = await convert(tableFixture(0xd660, concat(new Uint8Array([10]), modern())));
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  try {
    const xml = new TextDecoder().decode(archive.extract_image('word/document.xml'));
    expect(xml.indexOf('<w:shd ')).toBeLessThan(xml.indexOf('</w:tblPr>'));
    expect(xml.match(/<w:shd /g)).toHaveLength(1);
    expect(() => archive.parse()).not.toThrow();
  } finally { archive.free(); }
});

it.each([
  [0xd612, new Uint8Array([1, 0])],
  [0xd660, new Uint8Array([9, ...new Uint8Array(9)])],
  [0xd62d, concat(new Uint8Array([12, 0, 3]), modern())],
  [0xd609, new Uint8Array([2, 31, 0])],
])('rejects malformed shading operands through the public converter, code=%i', async (code, operand) => {
  await expect(convert(tableFixture(code as number, operand as Uint8Array)))
    .rejects.toMatchObject({ reason: 'unsupported-input', from: 'doc', to: 'docx' });
});

it('warns rather than inventing an OOXML pattern for unmappable binary coverage', async () => {
  const result = await convert(tableFixture(0xd612, concat(new Uint8Array([10]), modern(0x23))));
  expect(result.warnings?.some(w => w.includes('table'))).toBe(true);
  const archive = new DocxArchive(new Uint8Array(result.bytes));
  try { expect(new TextDecoder().decode(archive.extract_image('word/document.xml'))).not.toContain('<w:shd '); }
  finally { archive.free(); }
});
