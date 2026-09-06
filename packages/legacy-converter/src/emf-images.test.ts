import { readFile } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import { expect, it, vi } from 'vitest';
import initDocx, { DocxArchive } from '../../docx/src/wasm/docx_parser.js';
import initPptx, { PptxArchive } from '../../pptx/src/wasm/pptx_parser.js';
import { playEmf } from '../../core/src/image/emf.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildDocFixture, buildPptFixture, concat, little16, little32 } from './test-fixtures.js';

await initDocx({ module_or_path: await readFile(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url)) });
await initPptx({ module_or_path: await readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
const record = (kind: number, options: number, bytes: Uint8Array) => concat(little16(options), little16(kind), little32(bytes.length), bytes);

// MS-EMF: header, explicit stock black pen, rectangle, EOF. No external resources.
const emf = new Uint8Array(144);
const emfView = new DataView(emf.buffer);
for (const [offset, value] of [[0, 1], [4, 88], [16, 100], [20, 100], [32, 2540], [36, 2540],
  [40, 0x464d4520], [44, 0x10000], [48, 144], [52, 4], [56, 1], [72, 96], [76, 96], [80, 25], [84, 25],
  [88, 37], [92, 12], [96, 0x80000007],
  [100, 43], [104, 24], [108, 10], [112, 20], [116, 70], [120, 80], [124, 14], [128, 20], [140, 20]]) {
  emfView.setUint32(offset, value, true);
}
function blip(compressed: boolean, two: boolean, declared = emf.length): Uint8Array {
  const image = compressed ? new Uint8Array(deflateSync(emf)) : emf;
  const header = new Uint8Array(34);
  const view = new DataView(header.buffer);
  view.setUint32(0, declared, true);
  view.setUint32(28, image.length, true);
  header[32] = compressed ? 0 : 0xfe;
  header[33] = 0xfe;
  return record(0xf01a, (two ? 0x3d5 : 0x3d4) << 4, concat(new Uint8Array(two ? 32 : 16), header, image));
}
function doc(bytes: Uint8Array): Uint8Array {
  const picf = new Uint8Array(68);
  const view = new DataView(picf.buffer);
  for (const [offset, value] of [[4, 68], [6, 100], [28, 1440], [30, 1440], [32, 1000], [34, 1000]]) view.setUint16(offset, value, true);
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(1), little32(0x800))),
    record(0xf00b, 0x13, concat(little16(0x0104), little32(1))),
  ));
  const data = concat(picf, shape, bytes);
  new DataView(data.buffer).setUint32(0, data.length, true);
  return buildDocFixture({ text: '\u0001\u0001\r', data,
    characterProperties: concat(little16(0x0855), new Uint8Array([1]), little16(0x6a03), little32(0)) });
}
function ppt(bytes: Uint8Array, delayed: boolean): Uint8Array {
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(42), little32(0xa00))),
    record(0xf00b, 0x13, concat(little16(0x4104), little32(1))),
    record(0xf010, 0, concat(...[0, 0, 576, 576].map(little32))),
  ));
  const bse = new Uint8Array(36);
  const view = new DataView(bse.buffer);
  bse[0] = bse[1] = 2; // OfficeArtBlipEMF
  view.setUint32(20, bytes.length, true);
  view.setUint32(24, 1, true);
  const entry = record(0xf007, 0x22, concat(bse, delayed ? new Uint8Array() : bytes));
  return buildPptFixture(record(1036, 15, record(0xf002, 15, shape)), new Uint8Array(), undefined,
    { entries: [entry], ...(delayed ? { pictures: bytes } : {}) });
}
const cases = [false, true].flatMap(compressed => [false, true].flatMap(two => [
  { from: 'doc' as const, compressed, two, delayed: false },
  { from: 'ppt' as const, compressed, two, delayed: false },
  { from: 'ppt' as const, compressed, two, delayed: true },
]));
it.each(cases)('retains passive EMF through $from compressed=$compressed twoUIDs=$two delayed=$delayed', async ({ from, compressed, two, delayed }) => {
  const bytes = blip(compressed, two);
  const result = await converter.convert({ bytes: from === 'doc' ? doc(bytes) : ppt(bytes, delayed), from,
    to: from === 'doc' ? 'docx' : 'pptx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 });
  const archive = from === 'doc' ? new DocxArchive(new Uint8Array(result.bytes)) : new PptxArchive(new Uint8Array(result.bytes));
  try {
    const path = from === 'doc' ? 'word/media/image0.emf' : 'ppt/media/image1.emf';
    const extracted = archive.extract_image(path);
    expect(extracted).toEqual(emf);
    expect(new TextDecoder().decode(archive.extract_image('[Content_Types].xml'))).toContain('Extension="emf" ContentType="image/x-emf"');
    expect(new TextDecoder().decode(archive.parse())).toContain(path);
    const stroke = vi.fn();
    const context = new Proxy({ stroke }, { get(target, name) { return Reflect.get(target, name) ?? vi.fn(); } });
    expect(playEmf(extracted, context as unknown as CanvasRenderingContext2D, 100, 100)).toBe(true);
    expect(stroke).toHaveBeenCalled(); // Existing OOXML image player receives visible geometry.
  } finally { archive.free(); }
});

it.each(['doc', 'ppt'] as const)('rejects an EMF expansion-size bomb through the %s public converter', async from => {
  const bytes = blip(true, false, 0xffffffff);
  await expect(converter.convert({ bytes: from === 'doc' ? doc(bytes) : ppt(bytes, true), from,
    to: from === 'doc' ? 'docx' : 'pptx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 }))
    .rejects.toMatchObject({ reason: 'unsupported-input' });
});
