/// <reference types="vite/client" />
import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildPptFixture, concat, little16, little32 } from '../../legacy-converter/src/test-fixtures.ts';
import { materializePptxPresentation, openPptxPresentation } from './pptx.ts';
import { renderSlideNode, type NodeCanvasFactory } from './render.ts';
import { loadSkiaForTests } from './test-imports.ts';

const skia = await loadSkiaForTests();
const record = (options: number, kind: number, payload: Uint8Array) => concat(little16(options), little16(kind), little32(payload.length), payload);
const properties = (entries: [number, number][]) => record((entries.length << 4) | 3, 0xf00b, concat(...entries.map(([id, value]) => concat(little16(id), little32(value)))));
function atom(parent: number, flags: number) {
  return record(2, 1007, concat(new Uint8Array(12), little32(parent), new Uint8Array(4), little16(flags), little16(0)));
}
function shape(id: number, left: number, color: number, extra: Uint8Array = new Uint8Array()) {
  return record(15, 0xf004, concat(
    record(0x12, 0xf00a, concat(little32(id), little32(0xa00))),
    record(0, 0xf010, concat(little32(576), little32(left), little32(left + 1152), little32(1728))),
    properties([[0x181, color], [0x1ff, 0x00080000]]), extra,
  ));
}
function fixture(flags: number) {
  const drawing = (shapes: Uint8Array) => record(15, 1036, record(15, 0xf002, shapes));
  const master = concat(atom(0, 7), drawing(concat(
    shape(900, 576, 0x08000004),
    shape(901, 2304, 0xff00, record(15, 0xf011, record(0, 3011, concat(little32(0), new Uint8Array([1, 0, 0, 0]))))),
    shape(902, 4032, 0xff00, properties([[0x3bf, 0x00020002]])),
  )));
  const scheme = record(0x10, 2032, concat(...Array.from({ length: 8 }, () => little32(255))));
  return buildPptFixture(concat(atom(100, flags), scheme, drawing(shape(10, 1152, 0xff0000))), undefined, master);
}
async function convert(flags: number) {
  const wasm = await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));
  return materializePptxPresentation(fixture(flags), { legacyConversion: { ppt: { converter: createLegacyOfficeWasmConverter({ wasm }) } } });
}

it('inherits only enabled non-placeholder visible master objects before slide objects', async () => {
  const shown = await convert(1);
  expect(shown.slides[0].elements).toHaveLength(2);
  expect(shown.slides[0].elements).toMatchObject([
    { type: 'shape', fill: { color: 'FF0000' } },
    { type: 'shape', fill: { color: '0000FF' } },
  ]);
  const hidden = await convert(0);
  expect(hidden.slides[0].elements).toHaveLength(1);
});

it('replaces a master slide-number metacharacter without replacing literal asterisks', async () => {
  const textbox = record(15, 0xf00d, concat(record(0, 3999, little32(4)), record(0, 4008, new TextEncoder().encode('* / *')), record(0, 4056, little32(0))));
  const master = concat(atom(0, 7), record(15, 1036, record(15, 0xf002, shape(900, 576, 255, textbox))));
  const input = buildPptFixture(concat(atom(100, 1), record(15, 1036, record(15, 0xf002, new Uint8Array()))), undefined, master);
  const wasm = await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));
  const result = await materializePptxPresentation(input, { legacyConversion: { ppt: { converter: createLegacyOfficeWasmConverter({ wasm }) } } });
  // The synthetic DocumentAtom starts numbering at zero, which MS-PPT permits.
  const shapeModel = result.slides[0].elements[0];
  expect(shapeModel.type).toBe('shape');
  if (shapeModel.type !== 'shape') throw new Error('expected shape');
  expect(shapeModel.textBody?.paragraphs.flatMap(p => p.runs.map(r => 'text' in r ? r.text : '')).join('')).toBe('0 / *');
});

it('resolves slide-number positions from outline text through the ordinary parser', async () => {
  const outline = concat(record(0, 3999, little32(4)), record(0, 4008, new TextEncoder().encode('* / *')), record(0, 4056, little32(4)));
  const drawing = record(15, 1036, record(15, 0xf002, shape(10, 576, 255, record(15, 0xf00d, record(0, 3998, little32(0))))));
  const wasm = await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));
  const result = await materializePptxPresentation(buildPptFixture(drawing, outline), { legacyConversion: { ppt: { converter: createLegacyOfficeWasmConverter({ wasm }) } } });
  const shapeModel = result.slides[0].elements[0];
  if (shapeModel.type !== 'shape') throw new Error('expected shape');
  expect(shapeModel.textBody?.paragraphs.flatMap(p => p.runs.map(r => 'text' in r ? r.text : '')).join('')).toBe('* / 0');
});

it.skipIf(!skia)('paints inherited master graphics behind local shapes with the destination slide scheme', async () => {
  const { Canvas } = skia as typeof import('skia-canvas');
  const parsed = await convert(1);
  const canvas = new Canvas(960, 720);
  await renderSlideNode(canvas, parsed, 0, { width: 960, dpr: 1 });
  const pixel = (x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data);
  expect(pixel(140, 140)).toEqual([255, 0, 0, 255]);
  expect(pixel(240, 140)).toEqual([0, 0, 255, 255]);
  expect(pixel(450, 140)).toEqual([255, 255, 255, 255]);
  expect(pixel(730, 140)).toEqual([255, 255, 255, 255]);
});

it.skipIf(!skia)('resolves a passive master image through the same slide resource session', async () => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const source = new Canvas(8, 8);
  source.getContext('2d').fillStyle = '#00ff00';
  source.getContext('2d').fillRect(0, 0, 8, 8);
  const png = new Uint8Array(await source.toBuffer('png'));
  const picture = record(15, 0xf004, concat(
    record((75 << 4) | 2, 0xf00a, concat(little32(900), little32(0xa00))),
    record(0, 0xf010, concat(little32(576), little32(576), little32(1728), little32(1728))),
    properties([[0x4104, 1]]),
  ));
  const master = concat(atom(0, 7), record(15, 1036, record(15, 0xf002, picture)));
  const input = buildPptFixture(concat(atom(100, 1), record(15, 1036, record(15, 0xf002, new Uint8Array()))), undefined, master,
    { entries: [record(0x6e00, 0xf01e, concat(new Uint8Array(17), png))] });
  const wasm = await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));
  const session = await openPptxPresentation(input, { legacyConversion: { ppt: { converter: createLegacyOfficeWasmConverter({ wasm }) } } });
  const factory: NodeCanvasFactory = {
    createCanvas: (w, h) => new Canvas(w, h),
    loadImage: bytes => loadImage(Buffer.from(new Uint8Array(bytes))),
  };
  try {
    for await (const slide of session) {
      expect(slide.elements).toMatchObject([{ type: 'picture', imagePath: 'ppt/media/image1.png' }]);
      expect(new Uint8Array(await (await session.getImage('ppt/media/image1.png', 'image/png')).arrayBuffer())).toEqual(png);
      const canvas = new Canvas(960, 720);
      await session.renderSlide(canvas, slide, { width: 960, dpr: 1, factory });
      expect(Array.from(canvas.getContext('2d').getImageData(140, 140, 1, 1).data)).toEqual([0, 255, 0, 255]);
    }
  } finally { await session.close(); }
});
