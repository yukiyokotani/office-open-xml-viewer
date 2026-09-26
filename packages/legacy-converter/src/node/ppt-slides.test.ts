// Binary PowerPoint slide-level facts (visibility, master objects, backgrounds)
// through the direct reader and the ordinary PPTX model and Canvas renderer.
import { describe, expect, it } from 'vitest';
import { renderSlideNode, skia, skiaFactory } from './node-facade.js';
import {
  anchor, buildPptFixture, concat, drawing, little16, little32, materialize, openSession, properties, record,
  shapeAtom, slideAtom, spContainer, utf16le, zeroOriginRuler,
} from './ppt-records.js';

const pixelOf = (canvas: { getContext(kind: '2d'): { getImageData(x: number, y: number, w: number, h: number): { data: ArrayLike<number> } } }) =>
  (x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data);

describe('slide visibility', () => {
  // MS-PPT 2.6.6 SlideShowSlideInfoAtom.fHidden -> the model's Slide.hidden.
  const info = (flags: number) => record(0, 0x03f9, concat(new Uint8Array(10), little16(flags), new Uint8Array(4)));
  const content = drawing(spContainer(shapeAtom(202, 42, 0xa00), anchor(0, 0, 5760, 4320), record(15, 0xf00d, concat(
    record(0, 3999, little32(4)), record(0, 4000, utf16le('Retained hidden slide content')), zeroOriginRuler(),
  ))));
  const input = (payload: Uint8Array, master?: Uint8Array) => buildPptFixture(concat(payload, content), new Uint8Array(), master);

  it('uses only the fHidden bit and keeps the hidden slide content', async () => {
    // Every flag word is covered by ppt.rs::slide_visibility_uses_only_the_hidden_flag_in_all_flag_words.
    for (const flags of [0, 4, 0xfffb, 0xffff]) {
      const presentation = await materialize(input(info(flags)));
      expect(presentation.slides).toHaveLength(1);
      expect(presentation.slides[0].hidden ?? false).toBe((flags & 4) !== 0);
      expect(presentation.slides[0].elements).toMatchObject([
        { type: 'shape', textBody: { paragraphs: [{ runs: [{ text: 'Retained hidden slide content' }] }] } },
      ]);
    }
  });

  it('does not inherit master or nested visibility without a slide info atom', async () => {
    const presentation = await materialize(input(record(15, 5000, info(4)), info(4)));
    expect(presentation.slides[0].hidden ?? false).toBe(false);
  });

  it('rejects ambiguous or malformed visibility metadata', async () => {
    for (const payload of [
      concat(info(0), info(4)), record(1, 0x03f9, new Uint8Array(16)), record(16, 0x03f9, new Uint8Array(16)),
      record(0, 0x03f9, new Uint8Array(15)), record(0, 0x03f9, new Uint8Array(17)),
    ]) {
      await expect(materialize(input(payload))).rejects.toThrow(/SlideShowSlideInfoAtom/);
    }
  });
});

describe('master objects', () => {
  const shape = (id: number, left: number, color: number, extra: Uint8Array = new Uint8Array()) => spContainer(
    record(0x12, 0xf00a, concat(little32(id), little32(0xa00))), anchor(576, left, left + 1152, 1728),
    properties([[0x181, color], [0x1ff, 0x00080000]]), extra,
  );
  const input = (flags: number) => {
    const master = concat(slideAtom(0, 7), drawing(
      shape(900, 576, 0x08000004),
      // A placeholder and a hidden master shape are not inherited.
      shape(901, 2304, 0xff00, record(15, 0xf011, record(0, 3011, concat(little32(0), new Uint8Array([1, 0, 0, 0]))))),
      shape(902, 4032, 0xff00, properties([[0x3bf, 0x00020002]])),
    ));
    const scheme = record(0x10, 2032, concat(...Array.from({ length: 8 }, () => little32(255))));
    return buildPptFixture(concat(slideAtom(100, flags), scheme, drawing(shape(10, 1152, 0xff0000))), undefined, master);
  };

  it('inherits only enabled, visible non-placeholder master objects below slide objects', async () => {
    const shown = await materialize(input(1));
    expect(shown.slides[0].elementSources).toEqual([{ origin: 'master' }, { origin: 'slide' }]);
    // The master scheme reference resolves with the destination slide's scheme.
    expect(shown.slides[0].elements).toMatchObject([
      { type: 'shape', fill: { color: 'FF0000' } }, { type: 'shape', fill: { color: '0000FF' } },
    ]);
    expect((await materialize(input(0))).slides[0].elementSources).toEqual([{ origin: 'slide' }]);
    if (!skia) return;
    const canvas = new skia.Canvas(960, 720);
    await renderSlideNode(canvas as never, shown, 0, { width: 960, dpr: 1 });
    const pixel = pixelOf(canvas);
    expect(pixel(140, 140)).toEqual([255, 0, 0, 255]);
    expect(pixel(240, 140)).toEqual([0, 0, 255, 255]);
    expect(pixel(450, 140)).toEqual([255, 255, 255, 255]);
    expect(pixel(730, 140)).toEqual([255, 255, 255, 255]);
  });

  it.skipIf(!skia)('resolves a master picture through the slide resource session', async () => {
    const source = new skia!.Canvas(8, 8);
    source.getContext('2d').fillStyle = '#00ff00';
    source.getContext('2d').fillRect(0, 0, 8, 8);
    const png = new Uint8Array(await source.toBuffer('png'));
    const picture = spContainer(shapeAtom(75, 900, 0xa00), anchor(576, 576, 1728, 1728), properties([[0x4104, 1]]));
    const input = buildPptFixture(concat(slideAtom(100, 1), drawing()), undefined, concat(slideAtom(0, 7), drawing(picture)),
      { entries: [record(0x6e00, 0xf01e, concat(new Uint8Array(17), png))] });
    const session = await openSession(input);
    try {
      for await (const slide of session) {
        expect(slide.elementSources).toEqual([{ origin: 'master' }]);
        const [element] = slide.elements;
        if (element?.type !== 'picture') throw new Error('expected a master picture');
        expect(new Uint8Array(await (await session.getImage(element.imagePath, 'image/png')).arrayBuffer())).toEqual(png);
        const canvas = new skia!.Canvas(960, 720);
        await session.renderSlide(canvas as never, slide, { width: 960, dpr: 1, factory: skiaFactory() });
        expect(pixelOf(canvas)(140, 140)).toEqual([0, 255, 0, 255]);
      }
    } finally { await session.close(); }
  });
});

describe('backgrounds', () => {
  it('paints the followed master background with the destination scheme, or the slide background, behind foreground shapes', async () => {
    const shape = (flags: number, color: number) => spContainer(record(0x12, 0xf00a, concat(little32(1), little32(flags))),
      ...(flags & 0x400 ? [] : [anchor(576, 576, 1152, 1152)]), properties([[0x181, color]]));
    const master = drawing(shape(0xc00, 0x08000000));
    const slide = (flags: number) => concat(
      record(2, 1007, concat(new Uint8Array(12), little32(100), little32(0), little16(flags), little16(0))),
      record(0x10, 2032, concat(...Array.from({ length: 8 }, () => little32(0xff0000)))),
      drawing(shape(0xc00, 0x00ff00), shape(0xa00, 0x0000ff)),
    );
    for (const [flags, color, expected] of [[4, '0000FF', [0, 0, 255, 255]], [0, '00FF00', [0, 255, 0, 255]]] as const) {
      const presentation = await materialize(buildPptFixture(slide(flags), undefined, master));
      expect(presentation.slides[0].background).toMatchObject({ fillType: 'solid', color });
      expect(presentation.slides[0].elements).toMatchObject([{ fill: { fillType: 'solid', color: 'FF0000' } }]);
      if (!skia) continue;
      const canvas = new skia.Canvas(960, 720);
      await renderSlideNode(canvas as never, presentation, 0, { width: 960, dpr: 1 });
      expect(pixelOf(canvas)(10, 10)).toEqual(expected);
      expect(pixelOf(canvas)(120, 120)).toEqual([255, 0, 0, 255]);
    }
  });

  it.skipIf(!skia)('renders a stretched, half-transparent master background picture through the lazy image session', async () => {
    const raster = new skia!.Canvas(2, 2);
    raster.getContext('2d').fillStyle = '#00ff00';
    raster.getContext('2d').fillRect(0, 0, 2, 2);
    const png = new Uint8Array(await raster.toBuffer('png'));
    const master = drawing(spContainer(record(0x12, 0xf00a, concat(little32(1), little32(0xc00))),
      properties([[0x180, 3], [0x4186, 1], [0x182, 32768]])));
    // The direct reader requires the slide's own drawing (here empty).
    const slide = concat(record(2, 1007, concat(new Uint8Array(12), little32(100), little32(0), little16(4), little16(0))), drawing());
    const session = await openSession(buildPptFixture(slide, undefined, master, { entries: [record(0x6e00, 0xf01e, concat(new Uint8Array(17), png))] }));
    try {
      for await (const slide of session) {
        expect(slide.elements).toHaveLength(0);
        expect(slide.background).toMatchObject({ fillType: 'image', mimeType: 'image/png', stretch: true, alpha: 0.5 });
        const canvas = new skia!.Canvas(960, 720);
        await session.renderSlide(canvas as never, slide, { width: 960, dpr: 1, factory: skiaFactory() });
        for (const [x, y] of [[10, 10], [500, 600], [949, 709]]) {
          const pixel = pixelOf(canvas)(x, y);
          expect(pixel[1]).toBe(255);
          expect(pixel[0]).toBeGreaterThanOrEqual(127);
          expect(pixel[0]).toBeLessThanOrEqual(128);
          expect(pixel[2]).toBe(pixel[0]);
          expect(pixel[3]).toBe(255);
        }
      }
    } finally { await session.close(); }
  });

  it('rejects a slide without a drawing instead of projecting outline text', async () => {
    // Rust: ppt/drawing/direct_model.rs::rejects_missing_local_drawing_instead_of_silently_losing_outline_text.
    await expect(materialize(buildPptFixture())).rejects.toThrow(/no drawing/);
  });
});
