import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import initPptx, { PptxArchive } from '../../pptx/src/wasm/pptx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildPptFixture, concat, little16, little32 } from './test-fixtures.js';

await initPptx({ module_or_path: await readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
const record = (kind: number, bytes: Uint8Array, options = 0) => concat(little16(options), little16(kind), little32(bytes.length), bytes);
function input(ruler: Uint8Array, outline = false): Uint8Array {
  const text = new TextEncoder().encode('A\tB\rC\tD');
  const body = concat(record(3999, little32(4)), record(4008, text));
  const textbox = record(0xf00d, concat(outline ? record(3998, little32(0)) : body, record(4006, ruler)), 15);
  const shape = record(0xf004, concat(
    record(0xf00a, concat(little32(42), little32(0xa00)), (202 << 4) | 2),
    record(0xf010, concat(...[0, 0, 5760, 4320].map(little32))), textbox), 15);
  return buildPptFixture(record(1036, record(0xf002, shape, 15), 15), outline ? body : undefined);
}
const request = (ruler: Uint8Array, outline = false) => converter.convert({
  bytes: input(ruler, outline), from: 'ppt', to: 'pptx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024,
});
const stops = concat(little32(4), little16(4), ...[-576, 0, 576, 1152].map((p, i) => concat(little16(p & 65535), little16(i))));

it.each([false, true])('passes local ruler stops to the ordinary PPTX parser (outline=%s)', async outline => {
  const result = await request(stops, outline);
  const archive = new PptxArchive(new Uint8Array(result.bytes));
  try {
    const xml = new TextDecoder().decode(archive.extract_image('ppt/slides/slide1.xml'));
    expect(xml.match(/<a:tabLst>/g)).toHaveLength(2);
    const model = JSON.parse(new TextDecoder().decode(archive.parse()));
    for (const paragraph of model.slides[0].elements[0].textBody.paragraphs) {
      expect(paragraph.tabStops).toEqual([
        { pos: -914400, algn: 'l' }, { pos: 0, algn: 'ctr' },
        { pos: 914400, algn: 'r' }, { pos: 1828800, algn: 'dec' },
      ]);
    }
  } finally { archive.free(); }
});

it('keeps an explicitly empty list distinct from an absent local list in XML', async () => {
  for (const [ruler, expected] of [[little32(0), false], [concat(little32(4), little16(0)), true]] as const) {
    const result = await request(ruler);
    const archive = new PptxArchive(new Uint8Array(result.bytes));
    try {
      expect(new TextDecoder().decode(archive.extract_image('ppt/slides/slide1.xml')).includes('<a:tabLst>')).toBe(expected);
    } finally { archive.free(); }
  }
});

it('rejects truncated tab arrays and invalid alignment before producing an archive', async () => {
  const invalid = stops.slice(); new DataView(invalid.buffer).setUint16(8, 65535, true);
  for (const ruler of [stops.slice(0, -1), invalid, concat(stops, little16(0))]) {
    await expect(request(ruler)).rejects.toMatchObject({ reason: 'unsupported-input' });
  }
});
