import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initPptx, { PptxArchive } from '../../pptx/src/wasm/pptx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildPptFixture, concat, little16, little32, utf16le } from './test-fixtures.js';

await initPptx({ module_or_path: await readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
const record = (options: number, kind: number, bytes: Uint8Array) => concat(little16(options), little16(kind), little32(bytes.length), bytes);
const label = 'אבג ABC 123';
function fixture(direction: number, alignment = 0): Uint8Array {
  const style = concat(little32(label.length + 1), little16(0), little32(0x200800),
    little16(alignment), little16(direction), little32(label.length + 1), little32(0));
  const textbox = record(15, 0xf00d, concat(record(0, 3999, little32(0)),
    record(0, 4000, utf16le(label)), record(0, 4001, style)));
  const shape = record(15, 0xf004, concat(
    record((202 << 4) | 2, 0xf00a, concat(little32(42), little32(0xa00))),
    record(0, 0xf010, concat(...[0, 0, 2304, 1152].map(little32))), textbox));
  return buildPptFixture(record(15, 1036, record(15, 0xf002, shape)));
}
const convert = (direction: number, alignment?: number) => converter.convert({
  bytes: fixture(direction, alignment), from: 'ppt', to: 'pptx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024,
});

it.each([[0, 0], [1, 0], [0, 2], [1, 2]])('retains direction %i independently of explicit alignment %i', async (direction, alignment) => {
  const result = await convert(direction, alignment);
  const archive = new PptxArchive(new Uint8Array(result.bytes));
  try {
    const xml = new TextDecoder().decode(archive.extract_image('ppt/slides/slide1.xml'));
    expect(xml).toContain(`rtl="${direction}"`);
    expect(xml).toContain(`algn="${alignment === 0 ? 'l' : 'r'}"`);
    expect(xml).toContain(`<a:t>${label}</a:t>`); // Keep logical Unicode order.
    const model = JSON.parse(new TextDecoder().decode(archive.parse()));
    const paragraph = model.slides[0].elements[0].textBody.paragraphs[0];
    // The ordinary parser omits false from JSON; its default remains LTR.
    expect(paragraph.rtl ?? false).toBe(direction === 1);
    expect(paragraph.alignment).toBe(alignment === 0 ? 'l' : 'r');
  } finally { archive.free(); }
});

it.each([2, 255, 65535])('rejects the reserved direction value %i through the public converter', async direction => {
  await expect(convert(direction)).rejects.toMatchObject({ reason: 'unsupported-input' });
});
