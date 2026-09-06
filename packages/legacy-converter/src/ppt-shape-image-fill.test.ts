import { readFile } from 'node:fs/promises';
import { crc32, inflateSync } from 'node:zlib';
import { expect, it } from 'vitest';
import initPptx, { PptxArchive } from '../../pptx/src/wasm/pptx_parser.js';
import { createLegacyOfficeWasmConverter } from './index.js';
import { buildPptShapeImageFillFixture, shapeFillPng } from './ppt-shape-image-fill-fixture.js';

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
