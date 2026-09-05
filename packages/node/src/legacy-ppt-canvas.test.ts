/// <reference types="vite/client" />
// The source converter's WASM URL import is resolved by Vitest for this test.
import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildPptFixture, concat, little16, little32, utf16le } from '../../legacy-converter/src/test-fixtures.ts';
import { materializePptxPresentation, openPptxPresentation } from './pptx.ts';
import { renderSlideNode, type NodeCanvasFactory } from './render.ts';
import { loadSkiaForTests } from './test-imports.ts';

const skia = await loadSkiaForTests();
const converterWasm = readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));

describe('binary PowerPoint to ordinary Canvas rendering', () => {
  it.skipIf(!skia)('paints inherited scheme-colored backgrounds behind foreground shapes', async () => {
    const { Canvas } = skia as typeof import('skia-canvas');
    const record = (options: number, kind: number, payload: Uint8Array) => concat(little16(options), little16(kind), little32(payload.length), payload);
    const shape = (flags: number, color: number) => record(15, 0xf004, concat(
      record(0x12, 0xf00a, concat(little32(1), little32(flags))),
      ...(flags & 0x400 ? [] : [record(0, 0xf010, concat(little32(576), little32(576), little32(1152), little32(1152)))]),
      record(0x13, 0xf00b, concat(little16(0x181), little32(color))),
    ));
    const master = record(15, 1036, record(15, 0xf002, shape(0xc00, 0x08000000)));
    const slide = (flags: number) => concat(
      record(2, 1007, concat(new Uint8Array(12), little32(100), little32(0), little16(flags), little16(0))),
      record(0x10, 2032, concat(...Array.from({ length: 8 }, () => little32(0xff0000)))),
      record(15, 1036, record(15, 0xf002, concat(shape(0xc00, 0x00ff00), shape(0xa00, 0x0000ff)))),
    );
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    for (const [flags, expected] of [[4, [0, 0, 255, 255]], [0, [0, 255, 0, 255]]] as const) {
      const presentation = await materializePptxPresentation(buildPptFixture(slide(flags), undefined, master), { legacyConversion: { ppt: { converter } } });
      expect(presentation.slides[0].elements).toHaveLength(1);
      const canvas = new Canvas(960, 720);
      await renderSlideNode(canvas, presentation, 0, { width: 960, dpr: 1 });
      const pixel = (x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data);
      expect(pixel(10, 10)).toEqual(expected);
      expect(pixel(120, 120)).toEqual([255, 0, 0, 255]);
    }
  });
  it.skipIf(!skia)('renders a stretched master background image through the ordinary lazy image session', async () => {
    const { Canvas, loadImage } = skia as typeof import('skia-canvas');
    const record = (options: number, kind: number, payload: Uint8Array) => concat(little16(options), little16(kind), little32(payload.length), payload);
    const raster = new Canvas(2, 2);
    raster.getContext('2d').fillStyle = '#00ff00';
    raster.getContext('2d').fillRect(0, 0, 2, 2);
    const png = new Uint8Array(await raster.toBuffer('png'));
    const master = record(15, 1036, record(15, 0xf002, record(15, 0xf004, concat(
      record(0x12, 0xf00a, concat(little32(1), little32(0xc00))),
      record(0x33, 0xf00b, concat(little16(0x180), little32(3), little16(0x4186), little32(1), little16(0x182), little32(32768))),
    ))));
    const slide = record(2, 1007, concat(new Uint8Array(12), little32(100), little32(0), little16(4), little16(0)));
    const input = buildPptFixture(slide, undefined, master, { entries: [record(0x6e00, 0xf01e, concat(new Uint8Array(17), png))] });
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const session = await openPptxPresentation(input, { legacyConversion: { ppt: { converter } } });
    const factory: NodeCanvasFactory = {
      createCanvas: (w, h) => new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
      loadImage: (b) => loadImage(Buffer.from(new Uint8Array(b))) as unknown as ReturnType<NodeCanvasFactory['loadImage']>,
    };
    for await (const slide of session) {
      expect(slide.elements).toHaveLength(0);
      expect(slide.background).toMatchObject({ fillType: 'image', imagePath: 'ppt/media/image1.png' });
      const canvas = new Canvas(960, 720);
      await session.renderSlide(canvas, slide, { width: 960, dpr: 1, factory });
      for (const [x, y] of [[10, 10], [500, 600], [949, 709]]) {
        const pixel = Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data);
        expect(pixel[1]).toBe(255);
        expect(pixel[0]).toBeGreaterThanOrEqual(127);
        expect(pixel[0]).toBeLessThanOrEqual(128);
        expect(pixel[2]).toBe(pixel[0]);
        expect(pixel[3]).toBe(255);
      }
    }
  });
  it.skipIf(!skia)('renders embedded and delayed raster pictures with cropping and flips using ordinary PPTX images', async () => {
    const { Canvas } = skia as typeof import('skia-canvas');
    const record = (options: number, kind: number, payload: Uint8Array) => concat(little16(options), little16(kind), little32(payload.length), payload);
    const source = new Canvas(20, 10);
    const context = source.getContext('2d');
    context.fillStyle = '#ff0000'; context.fillRect(0, 0, 10, 10);
    context.fillStyle = '#0000ff'; context.fillRect(10, 0, 10, 10);
    const png = new Uint8Array(await source.toBuffer('png'));
    const jpeg = new Uint8Array(await source.toBuffer('jpg'));
    const pngBlip = record(0x6e00, 0xf01e, concat(new Uint8Array(17), png));
    const jpegBlip = record(0x46b0, 0xf01d, concat(new Uint8Array(33), jpeg));
    const entry = (type: number, size: number, offset: number, embedded: Uint8Array) => record((type << 4) | 2, 0xf007, concat(
      new Uint8Array([type, type]), new Uint8Array(18), little32(size), little32(1), little32(offset), new Uint8Array(4), embedded,
    ));
    const picture = (left: number, index: number, cropLeft = 0, flip = false) => record(15, 0xf004, concat(
      record((75 << 4) | 2, 0xf00a, concat(little32(42), little32(0xa00 | (flip ? 0x40 : 0)))),
      record(0, 0xf010, concat(little32(576), little32(left), little32(left + 1152), little32(1152))),
      record(0x23, 0xf00b, concat(little16(0x4104), little32(index), little16(0x102), little32(cropLeft))),
    ));
    const drawing = record(15, 1036, record(15, 0xf002, concat(picture(576, 1), picture(2304, 1, 32768), picture(4032, 2, 0, true))));
    const input = buildPptFixture(drawing, undefined, undefined, {
      entries: [entry(6, pngBlip.length, 0xffffffff, pngBlip), entry(5, jpegBlip.length, 23, new Uint8Array())],
      pictures: concat(new Uint8Array(23), jpegBlip),
    });
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const presentation = await materializePptxPresentation(input, { legacyConversion: { ppt: { converter } } });
    expect(presentation.slides[0].elements).toMatchObject([
      { type: 'picture', imagePath: 'ppt/media/image1.png' },
      { type: 'picture', imagePath: 'ppt/media/image1.png', srcRect: { l: 0.5, t: 0, r: 0, b: 0 } },
      { type: 'picture', imagePath: 'ppt/media/image2.jpg', flipH: true },
    ]);
    const canvas = new Canvas(960, 720);
    const factory: NodeCanvasFactory = {
      createCanvas: (w, h) => new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
      loadImage: (bytes) => (skia as typeof import('skia-canvas')).loadImage(Buffer.from(new Uint8Array(bytes))) as unknown as ReturnType<NodeCanvasFactory['loadImage']>,
    };
    const session = await openPptxPresentation(input, { legacyConversion: { ppt: { converter } } });
    try {
      expect(new Uint8Array(await (await session.getImage('ppt/media/image1.png', 'image/png')).arrayBuffer())).toEqual(png);
      await session.renderSlide(canvas, presentation.slides[0], { width: 960, dpr: 1, factory });
    } finally { await session.close(); }
    const pixel = (x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data);
    expect(pixel(120, 130)).toEqual([255, 0, 0, 255]);
    expect(pixel(260, 130)).toEqual([0, 0, 255, 255]);
    expect(pixel(420, 130)).toEqual([0, 0, 255, 255]);
    // JPEG is lossy; test dominant colors away from the source boundary.
    expect(pixel(700, 130)[2]).toBeGreaterThan(240);
    expect(pixel(830, 130)[0]).toBeGreaterThan(240);
  });
  it('inherits a main-master title size only for an actual placeholder, including outline text without direct styles', async () => {
    const record = (options: number, kind: number, payload: Uint8Array) => concat(little16(options), little16(kind), little32(payload.length), payload);
    const master = record(0, 4003, concat(little16(1), little32(0x800), little16(1), little32(0x20000), little16(48)));
    const outline = concat(record(0, 3999, little32(0)), record(0, 4000, utf16le('Title')));
    const shape = (position: number) => record(15, 0xf004, concat(
      record((202 << 4) | 2, 0xf00a, concat(little32(42), little32(0xa00))),
      record(0, 0xf010, concat(little32(576), little32(576), little32(3456), little32(1728))),
      record(15, 0xf00d, record(0, 3998, little32(0))),
      record(15, 0xf011, record(0, 3011, concat(little32(position), new Uint8Array([1, 0, 0, 0])))),
    ));
    const slide = concat(record(2, 1007, concat(new Uint8Array(12), little32(100), new Uint8Array(8))), record(15, 1036, record(15, 0xf002, concat(shape(0), shape(0xffffffff)))));
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const presentation = await materializePptxPresentation(buildPptFixture(slide, outline, master), { legacyConversion: { ppt: { converter } } });
    const [title, ordinary] = presentation.slides[0].elements;
    if (title.type !== 'shape' || ordinary.type !== 'shape') throw new Error('Expected text shapes');
    expect(title.textBody?.paragraphs[0]).toMatchObject({ alignment: 'ctr', runs: [{ text: 'Title', fontSize: 48 }] });
    expect(ordinary.textBody?.paragraphs[0].runs[0]).toMatchObject({ text: 'Title', fontSize: 18 });
  });
  it.skipIf(!skia)('paints nontext presets, solid lines and transparency through the ordinary renderer', async () => {
    const { Canvas } = skia as typeof import('skia-canvas');
    const record = (options: number, kind: number, payload: Uint8Array) => concat(little16(options), little16(kind), little32(payload.length), payload);
    const shape = (kind: number, left: number, color: number, opacity = 65536) => record(15, 0xf004, concat(
      record((kind << 4) | 2, 0xf00a, concat(little32(42), little32(0xa00))),
      record(0, 0xf010, concat(little32(576), little32(left), little32(left + 1152), little32(1728))),
      record((5 << 4) | 3, 0xf00b, concat(
        little16(0x181), little32(color), little16(0x182), little32(opacity),
        little16(0x1c0), little32(0xff0000), little16(0x1cb), little32(91440),
        little16(0x1ff), little32(0x00080008),
      )),
    ));
    const drawing = record(15, 1036, record(15, 0xf002, concat(
      shape(1, 576, 0xff), shape(3, 2304, 0xff, 32768), shape(20, 4032, 0xff),
    )));
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const presentation = await materializePptxPresentation(buildPptFixture(drawing), { legacyConversion: { ppt: { converter } } });
    expect(presentation.slides[0].elements).toHaveLength(3);
    const canvas = new Canvas(960, 720);
    await renderSlideNode(canvas, presentation, 0, { width: 960, dpr: 1 });
    const pixel = (x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data);
    expect(pixel(192, 192)).toEqual([255, 0, 0, 255]);
    // Half-opacity red ellipse over the presentation's white background.
    expect(pixel(480, 192)[0]).toBe(255);
    expect(pixel(480, 192)[1]).toBeGreaterThanOrEqual(127);
    expect(pixel(480, 192)[1]).toBeLessThanOrEqual(128);
    expect(pixel(390, 102)).toEqual([255, 255, 255, 255]);
    expect(pixel(96, 192)).toEqual([0, 0, 255, 255]);
    expect(pixel(768, 192)).toEqual([0, 0, 255, 255]);
    expect(pixel(720, 240)).toEqual([255, 255, 255, 255]);
  });
  it('retains outline-referenced size, alignment and local scheme colors in the ordinary parser model', async () => {
    const record = (options: number, kind: number, payload: Uint8Array) => concat(little16(options), little16(kind), little32(payload.length), payload);
    const outline = concat(
      record(0, 3999, little32(0)), record(0, 4000, utf16le('Wide')),
      record(0, 4001, concat(
        little32(5), little16(0), little32(0x800), little16(1),
        little32(5), little32(0x60000), little16(48), little32(0x05000000),
      )),
    );
    const drawing = record(15, 1036, record(15, 0xf002, record(15, 0xf004, concat(
      record((202 << 4) | 2, 0xf00a, concat(little32(42), little32(0x200))),
      record(0, 0xf010, concat(little32(576), little32(576), little32(3456), little32(1728))),
      record(15, 0xf00d, record(0, 3998, little32(0))),
    ))));
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const scheme = record(0x10, 2032, concat(...Array.from({ length: 8 }, () => little32(0x00563412))));
    const presentation = await materializePptxPresentation(buildPptFixture(concat(drawing, scheme), outline), { legacyConversion: { ppt: { converter } } });
    const element = presentation.slides[0].elements[0];
    expect(element.type).toBe('shape');
    if (element.type !== 'shape') throw new Error('Expected converted text shape');
    expect(element.textBody?.paragraphs[0]).toMatchObject({ alignment: 'ctr', runs: [{ text: 'Wide', fontSize: 48, color: '123456' }] });
  });
  it.skipIf(!skia)('renders two binary text frames at their own positions through the ordinary PPTX Canvas renderer', async () => {
    const { Canvas } = skia as typeof import('skia-canvas');
    const record = (options: number, kind: number, payload: Uint8Array) => concat(
      little16(options), little16(kind), little32(payload.length), payload,
    );
    const shape = (x: number, y: number, width: number, text: string) => record(15, 0xf004, concat(
      record((202 << 4) | 2, 0xf00a, concat(little32(42), little32(0x200))),
      // MS-PPT ClientAnchor is top, left, right, bottom in 1/576 inch.
      record(0, 0xf010, concat(little32(y), little32(x), little32(x + width), little32(y + 576))),
      record(15, 0xf00d, record(0, 4000, utf16le(text))),
    ));
    const drawing = record(15, 1036, record(15, 0xf002, concat(
      shape(576, 576, 1728, 'First'), shape(2880, 1728, 1152, 'Second'),
    )));
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const presentation = await materializePptxPresentation(buildPptFixture(drawing), {
      legacyConversion: { ppt: { converter } },
    });
    expect(presentation.slides[0].elements).toMatchObject([
      { x: 914400, y: 914400, width: 2743200, height: 914400 },
      { x: 4572000, y: 2743200, width: 1828800, height: 914400 },
    ]);
    const canvas = new Canvas(960, 720);
    await renderSlideNode(canvas, presentation, 0, { width: 960, dpr: 1 });
    const pixels = canvas.getContext('2d').getImageData(0, 0, 960, 720).data;
    const ink = (left: number, top: number, width: number, height: number) => {
      let count = 0;
      for (let y = top; y < top + height; y++) for (let x = left; x < left + width; x++) {
        const offset = (y * 960 + x) * 4;
        if (pixels[offset] < 128 && pixels[offset + 3] > 0) count++;
      }
      return count;
    };
    expect(ink(96, 96, 288, 96)).toBeGreaterThan(0);
    expect(ink(480, 288, 192, 96)).toBeGreaterThan(0);
    // The old single default-position text box painted here instead.
    expect(ink(48, 48, 400, 40)).toBe(0);
  });

});
