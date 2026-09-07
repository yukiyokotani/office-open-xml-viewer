import { readFile } from 'node:fs/promises';
import { crc32, inflateSync } from 'node:zlib';
import { expect, it } from 'vitest';
import initPptx, { PptxArchive } from '../../pptx/src/wasm/pptx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildPptShapeImageFillFixture, shapeFillPng } from './ppt-shape-image-fill-fixture.js';
import { concat, little16, little32 } from './test-fixtures.js';

await initPptx({ module_or_path: await readFile(new URL('../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url)) });
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('./wasm/legacy_office_converter_bg.wasm', import.meta.url)) });

it('converts a stretched foreground PPT shape image fill through the public pipeline', async () => {
  const png = Buffer.from(shapeFillPng);
  const idat = png.subarray(41, 55);
  expect(crc32(png.subarray(12, 29))).toBe(png.readUInt32BE(29));
  expect(crc32(png.subarray(37, 55))).toBe(png.readUInt32BE(55));
  expect(inflateSync(idat)).toEqual(Buffer.from([0, 255, 0, 0, 255, 0, 0, 255, 128]));
  const result = await converter.convert({
    bytes: buildPptShapeImageFillFixture(), from: 'ppt', to: 'pptx',
    signal: new AbortController().signal, maxOutputBytes: 1024 * 1024,
  });
  const archive = new PptxArchive(new Uint8Array(result.bytes));
  try {
    const slide = new TextDecoder().decode(archive.extract_image('ppt/slides/slide1.xml'));
    const rels = new TextDecoder().decode(archive.extract_image('ppt/slides/_rels/slide1.xml.rels'));
    expect(slide).toContain('<a:prstGeom prst="ellipse"');
    expect(slide).toContain('<a:blipFill rotWithShape="1"><a:blip r:embed="rImg1"><a:alphaModFix amt="50000"/></a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill>');
    expect(slide).toContain('<a:ln');
    expect(slide).toContain('<a:t>Picture fill text</a:t>');
    expect(rels).toContain('Id="rImg1"');
    expect(archive.extract_image('ppt/media/image1.png')).toEqual(shapeFillPng);
    const model = JSON.parse(new TextDecoder().decode(archive.parse()));
    expect(JSON.stringify(model.slides[0])).toContain('ppt/media/image1.png');
    expect(JSON.stringify(model.slides[0])).toContain('Picture fill text');
  } finally {
    archive.free();
  }
});

it('omits an unsupported WMF shape fill without dropping the shape text or leaving a relationship', async () => {
  const rec = (fn: number, words: number[] = []) => concat(little32(3 + words.length), little16(fn), ...words.map(little16));
  const body = concat(rec(0x020c, [100, 100]), rec(0));
  const wmf = concat(little16(1), little16(9), little16(0x300), little32((18 + body.length + 2) / 2), little16(0), little32(5), little16(0), body, new Uint8Array(2));
  const result = await converter.convert({ bytes: buildPptShapeImageFillFixture(wmf), from: 'ppt', to: 'pptx', signal: new AbortController().signal, maxOutputBytes: 1024 * 1024 });
  const archive = new PptxArchive(new Uint8Array(result.bytes));
  try {
    const slide = new TextDecoder().decode(archive.extract_image('ppt/slides/slide1.xml'));
    const rels = new TextDecoder().decode(archive.extract_image('ppt/slides/_rels/slide1.xml.rels'));
    expect(slide).toContain('<a:t>Picture fill text</a:t>');
    expect(slide).not.toContain('<a:blipFill');
    expect(rels).not.toContain('rImg1');
    expect(() => archive.extract_image('ppt/media/image1.wmf')).toThrow();
    expect(result.warnings).toContain('legacy-ppt:custom-geometry-unlinked-and-advanced-paint-unsupported-media-and-actions-omitted');
  } finally { archive.free(); }
});
