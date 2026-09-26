// Binary PowerPoint geometry, paint and pictures through the direct reader and
// the ordinary PPTX model and Canvas renderer. Property-level matrices live in
// the Rust unit tests (ppt/paint.rs, ppt/drawing.rs, officeart/preset.rs).
import { describe, expect, it } from 'vitest';
import { buildPptShapeImageFillFixture, shapeFillPng } from '../ppt-shape-image-fill-fixture.js';
import { renderSlideNode, skia, skiaFactory } from './node-facade.js';
import {
  anchor, buildPptFixture, concat, drawing, little16, little32, materialize, openSession, properties, record,
  shapeAtom, slideAtom, spContainer,
} from './ppt-records.js';

type Presentation = Awaited<ReturnType<typeof materialize>>;
type Rgba = [number, number, number, number];
const BLUE: Rgba = [0, 0, 255, 255], RED: Rgba = [255, 0, 0, 255], WHITE: Rgba = [255, 255, 255, 255], GREEN: Rgba = [0, 255, 0, 255];

/** One shape with the given OfficeArt properties (colours are 0x00BBGGRR: 0xff0000 is blue). */
const lineShape = (kind: number, place: Uint8Array, entries: [number, number][], flags = 0xa00) =>
  drawing(spContainer(shapeAtom(kind, 10, flags), place, properties(entries)));

describe('lines and connectors', () => {
  const connector = (place: Uint8Array, flags = 0xb00) =>
    materialize(buildPptFixture(lineShape(32, place, [[0x181, 255], [0x1c0, 0xff0000], [0x1cb, 38100], [0x1d1, 1]], flags)));

  it('keeps a straight connector with a zero-sized axis, its line and arrow, and no fill', async () => {
    for (const [place, width, height] of [
      [anchor(576, 576, 1728, 576), 1828800, 0], [anchor(576, 576, 576, 1728), 0, 1828800],
    ] as const) {
      const model = await connector(place);
      expect(model.slides[0].elements).toHaveLength(1);
      expect(model.slides[0].elements[0]).toMatchObject({
        type: 'shape', geometry: 'straightConnector1', width, height, fill: { fillType: 'none' },
        stroke: { color: '0000FF', tailEnd: { type: 'triangle' } },
      });
    }
  });

  it.skipIf(!skia)('paints the connector arrow at the authored end, including horizontal reflection', async () => {
    for (const flipped of [false, true]) {
      const pixel = await render(await connector(anchor(576, 576, 1728, 576), flipped ? 0xb40 : 0xb00));
      expect(pixel(192, 96)).toEqual(BLUE);
      expect(pixel(flipped ? 108 : 276, 100)).toEqual(BLUE);
      expect(pixel(flipped ? 276 : 108, 100)).toEqual(WHITE);
    }
  });

  const dashed = (dash: number) => materialize(buildPptFixture(
    lineShape(20, anchor(576, 576, 1728, 576), [[0x1c0, 0xff0000], [0x1cb, 38100], [0x1ce, dash], [0x1d1, 1]])));

  it('maps binary dash presets and rejects an undefined one', async () => {
    // Every preset is covered by ppt/paint.rs::all_dash_presets_retain_the_line_and_inherit_explicit_solid;
    // here: solid, the first system and the last preset reach the model with the arrow kept.
    for (const [id, name] of [[0, undefined], [1, 'sysDash'], [6, 'dash'], [10, 'lgDashDotDot']] as const) {
      const [shape] = (await dashed(id)).slides[0].elements;
      if (shape.type !== 'shape') throw new Error('expected a line shape');
      expect(shape.stroke).toMatchObject({ color: '0000FF', tailEnd: { type: 'triangle' } });
      expect(shape.stroke?.dashStyle).toBe(name);
    }
    await expect(dashed(11)).rejects.toThrow(/line dashing/);
  });

  it.skipIf(!skia)('paints a dashed leader with gaps and a solid arrow tip', async () => {
    const presentation = await dashed(6);
    const canvas = await canvasOf(presentation);
    const row = canvas.getContext('2d').getImageData(100, 96, 150, 1).data;
    let blue = 0, white = 0;
    for (let i = 0; i < row.length; i += 4) {
      if (row[i] === 0 && row[i + 1] === 0 && row[i + 2] === 255) blue++;
      if (row[i] === 255 && row[i + 1] === 255 && row[i + 2] === 255) white++;
    }
    expect(blue).toBeGreaterThan(0);
    expect(white).toBeGreaterThan(0);
    expect(Array.from(canvas.getContext('2d').getImageData(276, 100, 1, 1).data)).toEqual(BLUE);
  });

  const stroked = (entries: [number, number][]) => materialize(buildPptFixture(
    lineShape(20, anchor(576, 576, 1728, 576), [[0x1c0, 0xff0000], [0x1cb, 114300], ...entries])));

  it('keeps line-end sizes, cap, join and miter limit', async () => {
    const model = await stroked([[0x1cc, 0x18000], [0x1d0, 1], [0x1d1, 5], [0x1d2, 0], [0x1d3, 2], [0x1d4, 2], [0x1d5, 0], [0x1d6, 1], [0x1d7, 1]]);
    expect(model.slides[0].elements[0]).toMatchObject({ type: 'shape', stroke: {
      color: '0000FF', width: 114300, lineCap: 'square', lineJoin: 'miter', miterLimit: 1.5,
      headEnd: { type: 'triangle', w: 'sm', len: 'lg' }, tailEnd: { type: 'arrow', w: 'lg', len: 'sm' },
    } });
  });

  it.skipIf(!skia)('paints a round cap beyond the endpoint without moving the line anchor', async () => {
    const pixels = [];
    for (const cap of [0, 2]) pixels.push((await render(await stroked([[0x1d7, cap]])))(92, 96));
    expect(pixels).toEqual([BLUE, WHITE]);
  });
});

describe('preset geometry and paint', () => {
  it('paints nontext presets, solid lines and transparency', async () => {
    const shape = (kind: number, left: number, opacity = 65536) => spContainer(shapeAtom(kind, 42, 0xa00), anchor(576, left, left + 1152, 1728),
      properties([[0x181, 0xff], [0x182, opacity], [0x1c0, 0xff0000], [0x1cb, 91440], [0x1ff, 0x00080008]]));
    const presentation = await materialize(buildPptFixture(drawing(shape(1, 576), shape(3, 2304, 32768), shape(20, 4032))));
    expect(presentation.slides[0].elements).toMatchObject([
      { geometry: 'rect', fill: { fillType: 'solid', color: 'FF0000' }, stroke: { color: '0000FF', width: 91440 } },
      { geometry: 'ellipse', fill: { fillType: 'solid', color: 'FF000080' } },
      { geometry: 'line', fill: { fillType: 'none' }, stroke: { color: '0000FF' } },
    ]);
    if (!skia) return;
    const pixel = await render(presentation);
    expect(pixel(192, 192)).toEqual(RED);
    // Half-opacity red ellipse over the white slide.
    expect(pixel(480, 192)[0]).toBe(255);
    expect(pixel(480, 192)[1]).toBeGreaterThanOrEqual(127);
    expect(pixel(480, 192)[1]).toBeLessThanOrEqual(128);
    expect(pixel(390, 102)).toEqual(WHITE);
    expect(pixel(96, 192)).toEqual(BLUE);
    expect(pixel(768, 192)).toEqual(BLUE);
    expect(pixel(720, 240)).toEqual(WHITE);
  });

  it('inherits master paint only through active links and honours direct colour and no-paint overrides', async () => {
    const master = drawing(spContainer(record(0x12, 0xf00a, concat(little32(900), little32(0x800))),
      properties([[0x181, 255], [0x1c0, 0xff0000], [0x1cb, 91440]])));
    const shape = (index: number, direct: [number, number][], active = true) => {
      const left = 576 + index * 864;
      return spContainer(record(0x12, 0xf00a, concat(little32(index + 1), little32(active ? 0xa20 : 0xa00))),
        anchor(576, left, left + 576, 1152), properties([[0x301, active ? 900 : 999], ...direct]));
    };
    const slide = drawing(shape(0, [[0x1bf, 0]]), shape(1, [[0x181, 0xff00]]),
      shape(2, [[0x1bf, 0x00100000], [0x1ff, 0x00080000]]), shape(3, [], false));
    const presentation = await materialize(buildPptFixture(slide, undefined, master));
    expect(presentation.slides[0].elements).toMatchObject([
      { fill: { fillType: 'solid', color: 'FF0000' }, stroke: { color: '0000FF', width: 91440 } },
      { fill: { fillType: 'solid', color: '00FF00' }, stroke: { color: '0000FF' } },
      { fill: { fillType: 'none' }, stroke: null },
      { fill: { fillType: 'none' }, stroke: null },
    ]);
    if (!skia) return;
    const pixel = await render(presentation);
    expect(pixel(140, 140)).toEqual(RED);
    expect(pixel(284, 140)).toEqual(GREEN);
    expect(pixel(428, 140)).toEqual(WHITE);
    expect(pixel(572, 140)).toEqual(WHITE);
    expect(pixel(96, 140)).toEqual(BLUE);
    expect(pixel(384, 140)).toEqual(WHITE);
  });

  describe('custom cubic paths', () => {
    const array = (size: number, items: Uint8Array[]) => concat(little16(items.length), little16(items.length), little16(size), ...items);
    const vertices = array(8, [[10, 120], [10, 20], [110, 20], [110, 120], [10, 120]].map(([x, y]) => concat(little32(x), little32(y))));
    const segments = array(2, [0x4000, 0x2001, 1, 0x6001, 0x8000].map(little16));
    const place = anchor(576, 576, 1728, 1728);
    const shape = (id: number, kind: number, props: Uint8Array, flags = 0xa00) => spContainer(shapeAtom(kind, id, flags), place, props);
    const geometry: [number, number][] = [[0x140, 10], [0x141, 20], [0x142, 110], [0x143, 120], [0x144, 4], [0xc145, vertices.length], [0xc146, segments.length]];
    const input = (inherited: boolean) => {
      const master = concat(slideAtom(0, 1), drawing(shape(900, 1, properties([[0x181, 255], [0x1ff, 0x00080000]])),
        ...(inherited ? [shape(901, 0, properties([...geometry, [0x3bf, 0x00020002]], concat(vertices, segments)))] : [])));
      const front = inherited
        ? shape(10, 0, properties([[0x301, 901], [0x181, 0xff0000], [0x1ff, 0x00080000]]), 0xa20)
        : shape(10, 0, properties([...geometry, [0x181, 0xff0000], [0x1ff, 0x00080000]], concat(vertices, segments)));
      return buildPptFixture(concat(slideAtom(100, 1), drawing(front)), undefined, master);
    };

    it.each([false, true])('projects an explicit or master-linked path as custom geometry (inherited=%s)', async inherited => {
      const model = await materialize(input(inherited));
      expect(model.slides[0].elementSources).toEqual([{ origin: 'master' }, { origin: 'slide' }]);
      expect(model.slides[0].elements[1]).toMatchObject({ type: 'shape', geometry: 'custGeom', fill: { color: '0000FF' } });
    });

    it.skipIf(!skia)('occludes master objects only inside the foreground curve', async () => {
      const pixel = await render(await materialize(input(false)));
      expect(pixel(192, 240)).toEqual(BLUE);
      expect(pixel(110, 110)).toEqual(RED);
      expect(pixel(320, 240)).toEqual(WHITE);
    });
  });
});

describe('pictures and picture fills', () => {
  it.skipIf(!skia)('renders embedded and delayed raster pictures with cropping and flips', async () => {
    const source = new skia!.Canvas(20, 10);
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
    const picture = (left: number, index: number, cropLeft = 0, flip = false) => spContainer(
      shapeAtom(75, 42 + left, 0xa00 | (flip ? 0x40 : 0)), anchor(576, left, left + 1152, 1152), properties([[0x4104, index], [0x102, cropLeft]]));
    const input = buildPptFixture(drawing(picture(576, 1), picture(2304, 1, 32768), picture(4032, 2, 0, true)), undefined, undefined, {
      entries: [entry(6, pngBlip.length, 0xffffffff, pngBlip), entry(5, jpegBlip.length, 23, new Uint8Array())],
      pictures: concat(new Uint8Array(23), jpegBlip),
    });
    const session = await openSession(input);
    const canvas = new skia!.Canvas(960, 720);
    try {
      for await (const slide of session) {
        expect(slide.elements).toMatchObject([
          { type: 'picture', flipH: false }, { type: 'picture', srcRect: { l: 0.5, t: 0, r: 0, b: 0 } }, { type: 'picture', flipH: true },
        ]);
        const [first, second, third] = slide.elements.map(element => (element.type === 'picture' ? element.imagePath : ''));
        // One store entry referenced twice is one image; the delayed JPEG is another.
        expect(second).toBe(first);
        expect(third).not.toBe(first);
        expect(new Uint8Array(await (await session.getImage(first, 'image/png')).arrayBuffer())).toEqual(png);
        await session.renderSlide(canvas as never, slide, { width: 960, dpr: 1, factory: skiaFactory() });
      }
    } finally { await session.close(); }
    const pixel = (x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data);
    expect(pixel(120, 130)).toEqual(RED);
    expect(pixel(260, 130)).toEqual(BLUE);
    expect(pixel(420, 130)).toEqual(BLUE);
    // JPEG is lossy; test dominant colours away from the source boundary.
    expect(pixel(700, 130)[2]).toBeGreaterThan(240);
    expect(pixel(830, 130)[0]).toBeGreaterThan(240);
  });

  it('keeps a stretched shape picture fill with its opacity, line, geometry and text', async () => {
    const session = await openSession(buildPptShapeImageFillFixture());
    try {
      for await (const slide of session) {
        expect(slide.elements).toMatchObject([{
          type: 'shape', geometry: 'ellipse',
          fill: { fillType: 'image', mimeType: 'image/png', rotWithShape: true, stretch: true, alpha: 0.5 },
          stroke: { color: 'FF0000' },
          textBody: { paragraphs: [{ runs: [{ text: 'Picture fill text' }] }] },
        }]);
        const [shape] = slide.elements;
        if (shape.type !== 'shape' || shape.fill?.fillType !== 'image') throw new Error('expected an image fill');
        expect(new Uint8Array(await (await session.getImage(shape.fill.imagePath, 'image/png')).arrayBuffer())).toEqual(shapeFillPng);
      }
    } finally { await session.close(); }
  });

  it('rejects an unsupported metafile shape fill instead of silently dropping it', async () => {
    const rec = (fn: number, words: number[] = []) => concat(little32(3 + words.length), little16(fn), ...words.map(little16));
    const body = concat(rec(0x020c, [100, 100]), rec(0));
    const wmf = concat(little16(1), little16(9), little16(0x300), little32((18 + body.length + 2) / 2), little16(0), little32(5), little16(0), body, new Uint8Array(2));
    await expect(materialize(buildPptShapeImageFillFixture(wmf))).rejects.toThrow(/fill BLIP is not a supported image/);
  });
});

async function canvasOf(presentation: Presentation) {
  const canvas = new skia!.Canvas(960, 720);
  await renderSlideNode(canvas as never, presentation, 0, { width: 960, dpr: 1 });
  return canvas;
}

async function render(presentation: Presentation): Promise<(x: number, y: number) => number[]> {
  const context = (await canvasOf(presentation)).getContext('2d');
  return (x, y) => Array.from(context.getImageData(x, y, 1, 1).data);
}
