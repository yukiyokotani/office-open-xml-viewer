// Binary PowerPoint text through the direct reader and the ordinary PPTX model
// and Canvas renderer. Parser-level matrices (every numbering scheme, every
// style mask boundary, ruler encodings) live in the Rust unit tests under
// packages/legacy-converter/parser/src/ppt; these cases pin the public wiring.
import { describe, expect, it } from 'vitest';
import { renderSlideNode, skia } from './node-facade.js';
import {
  anchor, buildPptFixture, concat, drawing, little16, little32, materialize, properties, record, shapeAtom,
  spContainer, utf16le, zeroOriginRuler,
} from './ppt-records.js';

type Presentation = Awaited<ReturnType<typeof materialize>>;

const ascii = (text: string) => record(0, 4008, new TextEncoder().encode(text));
const header = (type: number) => record(0, 3999, little32(type));
const outlineReference = () => record(0, 3998, little32(0));
const textbox = (...parts: Uint8Array[]) => record(15, 0xf00d, concat(...parts));
const textShape = (box: Uint8Array, place = anchor(0, 0, 5760, 4320), ...extra: Uint8Array[]) =>
  spContainer(shapeAtom(202, 42, 0xa00), place, ...extra, box);
/** A slide whose text sits inline in the text box, or in the outline behind an OutlineTextRefAtom. */
const textSlide = (body: Uint8Array, outline = false, ruler = zeroOriginRuler()) => buildPptFixture(
  drawing(textShape(textbox(outline ? outlineReference() : body, ruler))), outline ? body : undefined,
);
const paragraphs = (presentation: Presentation, element = 0) => {
  const shape = presentation.slides[0].elements[element];
  if (shape?.type !== 'shape' || !shape.textBody) throw new Error('expected a text shape');
  return shape.textBody.paragraphs;
};
const textOf = (presentation: Presentation, element = 0) =>
  paragraphs(presentation, element).map(p => p.runs.map(r => ('text' in r ? r.text : '')).join('')).join('|');

describe('local PP9 automatic numbering', () => {
  const utf16 = (text: string) => concat(...Array.from(text, c => little16(c.charCodeAt(0))));
  const extension = (scheme = 3, start = 1, enabled = 1) => concat(
    little32(0x03800000), little16(65535), little16(enabled), little16(scheme), little16(start), little32(0), little32(0),
  );
  function input(ext: Uint8Array, groups = [[4, 0]], tag = '___PPT9', flags = 1, owner: 'inline' | 'outline' | 'no-header' = 'inline') {
    const body = concat(owner === 'no-header' ? new Uint8Array() : header(4), ascii('A\rB'),
      record(0, 4001, concat(little32(4), little16(0), little32(0x81), little16(flags), little16(0x2022),
        ...groups.map(([count, id]) => concat(little32(count), little32(0x400), little16(id << 10))))));
    const tags = record(15, 5000, record(15, 5002, concat(record(0, 4026, utf16(tag)), record(0, 5003, record(0, 4012, ext)))));
    const box = textbox(owner === 'outline' ? outlineReference() : body, zeroOriginRuler());
    return buildPptFixture(drawing(textShape(box, anchor(0, 0, 5760, 4320), record(15, 0xf011, tags))), owner === 'outline' ? body : undefined);
  }
  const bullets = async (bytes: Uint8Array) => paragraphs(await materialize(bytes)).map(p => p.bullet);

  it('applies explicit shape-local numbering to every paragraph and binds run groups in order', async () => {
    const numbered = { type: 'autoNum', numType: 'arabicPeriod', startAt: 6 };
    expect(await bullets(input(extension(3, 6)))).toMatchObject([numbered, numbered]);
    // The first entry (scheme 0, group 0) matches no run; later groups bind in order.
    expect(await bullets(input(concat(extension(0, 9), extension(3, 2), extension(7, 4)), [[2, 1], [2, 2]]))).toMatchObject([
      { type: 'autoNum', numType: 'arabicPeriod', startAt: 2 },
      { type: 'autoNum', numType: 'romanUcPeriod', startAt: 4 },
    ]);
    // A paragraph whose runs disagree keeps its character bullet.
    expect(await bullets(input(concat(extension(3, 1), extension(3, 2)), [[1, 0], [3, 1]]))).toMatchObject([
      { type: 'char' }, { type: 'autoNum', startAt: 2 },
    ]);
    // The last scheme of the table at the largest start reaches the model.
    expect((await bullets(input(extension(40, 32767))))[0]).toMatchObject({ type: 'autoNum', numType: 'hindiAlpha1Period', startAt: 32767 });
  });

  it('does not number foreign tags, unmatched runs, disabled entries, unbulleted text or text the shape does not own', async () => {
    for (const bytes of [
      input(extension(), [[4, 0]], 'notPPT9'), input(extension(), [[4, 1]]), input(extension(3, 1, 0)),
      input(extension(), [[4, 0]], '___PPT9', 1, 'outline'), input(extension(), [[4, 0]], '___PPT9', 1, 'no-header'),
    ]) {
      expect(await bullets(bytes)).toMatchObject([{ type: 'char' }, { type: 'char' }]);
    }
    expect(await bullets(input(extension(), [[4, 0]], '___PPT9', 0))).toMatchObject([{ type: 'none' }, { type: 'none' }]);
  });

  it('rejects malformed local numbering instead of dropping it', async () => {
    await expect(materialize(input(extension(41)))).rejects.toThrow(/automatic numbering scheme/);
    await expect(materialize(input(extension(3, 0)))).rejects.toThrow(/automatic numbering start/);
    await expect(materialize(input(extension().slice(0, -1)))).rejects.toThrow(/UNSUPPORTED/);
  });
});

it.each([false, true])('validates the signed character position domain without inventing a baseline (outline=%s)', async outline => {
  const body = (position: number) => concat(header(4), ascii('X'),
    record(0, 4001, concat(little32(2), little16(0), little32(0), little32(2), little32(0x80000), little16(position & 65535))));
  for (const position of [-100, 100]) expect(textOf(await materialize(textSlide(body(position), outline)))).toBe('X');
  for (const position of [-101, 101]) {
    await expect(materialize(textSlide(body(position), outline))).rejects.toThrow(/character baseline position/);
  }
});

describe('local ruler tabs', () => {
  const stops = concat(little16(4), ...[-576, 0, 576, 1152].map((p, i) => concat(little16(p & 65535), little16(i))));
  const body = concat(header(4), ascii('A\tB\rC\tD'));
  const ruler = (tabs: Uint8Array) => record(0, 4006, concat(little32(4 | 8 | 256), tabs, little32(0)));

  it.each([false, true])('applies the local stops to every paragraph in binary order (outline=%s)', async outline => {
    const presentation = await materialize(textSlide(body, outline, ruler(stops)));
    const expected = [{ pos: -914400, algn: 'l' }, { pos: 0, algn: 'ctr' }, { pos: 914400, algn: 'r' }, { pos: 1828800, algn: 'dec' }];
    expect(paragraphs(presentation).map(p => p.tabStops)).toEqual([expected, expected]);
  });

  it('rejects truncated tab arrays and invalid alignments', async () => {
    const invalid = stops.slice();
    new DataView(invalid.buffer).setUint16(4, 65535, true);
    await expect(materialize(textSlide(body, false, ruler(invalid)))).rejects.toThrow(/ruler tab alignment/);
    await expect(materialize(textSlide(body, false, record(0, 4006, concat(little32(4), stops.slice(0, -1)))))).rejects.toThrow(/ruler tabs/);
  });
});

describe('paragraph text direction', () => {
  const label = 'אבג ABC 123';
  const input = (direction: number, alignment = 0) => buildPptFixture(drawing(textShape(textbox(
    header(0), record(0, 4000, utf16le(label)),
    record(0, 4001, concat(little32(label.length + 1), little16(0), little32(0x200800), little16(alignment), little16(direction), little32(label.length + 1), little32(0))),
    zeroOriginRuler(),
  ), anchor(0, 0, 2304, 1152))));

  it('keeps direction independent of explicit alignment and the text in logical order', async () => {
    for (const [direction, alignment] of [[0, 0], [1, 0], [0, 2], [1, 2]]) {
      const [paragraph] = paragraphs(await materialize(input(direction, alignment)));
      expect(paragraph.rtl ?? false).toBe(direction === 1);
      expect(paragraph.alignment).toBe(alignment === 0 ? 'l' : 'r');
      expect(paragraph.runs.map(r => ('text' in r ? r.text : '')).join('')).toBe(label);
    }
  });

  it('rejects reserved direction values', async () => {
    for (const direction of [2, 65535]) await expect(materialize(input(direction))).rejects.toThrow(/text direction/);
  });
});

describe('text inheritance and placement', () => {
  it('resolves outline-referenced size, alignment and local scheme colours', async () => {
    const outline = concat(header(0), record(0, 4000, utf16le('Wide')), record(0, 4001, concat(
      little32(5), little16(0), little32(0x800), little16(1), little32(5), little32(0x60000), little16(48), little32(0x05000000),
    )));
    const scheme = record(0x10, 2032, concat(...Array.from({ length: 8 }, () => little32(0x00563412))));
    const slide = concat(drawing(textShape(textbox(outlineReference(), zeroOriginRuler()), anchor(576, 576, 3456, 1728))), scheme);
    expect(paragraphs(await materialize(buildPptFixture(slide, outline)))[0])
      .toMatchObject({ alignment: 'ctr', runs: [{ text: 'Wide', fontSize: 48, color: '123456' }] });
  });

  it('applies the main-master title style to title-typed outline text', async () => {
    // The direct model styles typed text from the master whether or not it is
    // a placeholder (Rust: typed_master_style_follows_text_type_not_only_placeholder_status).
    const master = record(0, 4003, concat(little16(1), little32(0x800), little16(1), little32(0x20000), little16(48)));
    const outline = concat(header(0), record(0, 4000, utf16le('Title')));
    const shape = spContainer(shapeAtom(202, 42, 0xa00), anchor(576, 576, 3456, 1728), textbox(outlineReference(), zeroOriginRuler()),
      record(15, 0xf011, record(0, 3011, concat(little32(0), new Uint8Array([1, 0, 0, 0])))));
    const slide = concat(record(2, 1007, concat(new Uint8Array(12), little32(100), new Uint8Array(8))), drawing(shape));
    expect(paragraphs(await materialize(buildPptFixture(slide, outline, master)))[0])
      .toMatchObject({ alignment: 'ctr', runs: [{ text: 'Title', fontSize: 48 }] });
  });

  it('substitutes slide-number metacharacters at their positions and leaves literal asterisks', async () => {
    const slideAtom = (master: number, flags: number) => record(2, 1007, concat(new Uint8Array(12), little32(master), new Uint8Array(4), little16(flags), little16(0)));
    const numbered = (position: number) => concat(header(4), ascii('* / *'), record(0, 4056, little32(position)));
    const frame = (box: Uint8Array) => spContainer(record(0x12, 0xf00a, concat(little32(900), little32(0xa00))), anchor(576, 576, 1728, 1728), box);
    // The synthetic DocumentAtom starts numbering at zero, which MS-PPT permits.
    const fromMaster = buildPptFixture(concat(slideAtom(100, 1), drawing()), undefined,
      concat(slideAtom(0, 7), drawing(frame(textbox(numbered(0), zeroOriginRuler())))));
    expect(textOf(await materialize(fromMaster))).toBe('0 / *');
    const fromOutline = buildPptFixture(drawing(frame(textbox(outlineReference(), zeroOriginRuler()))), numbered(4));
    expect(textOf(await materialize(fromOutline))).toBe('* / 0');
  });

  it('inherits text style only through an active, resolvable hspMaster link and keeps direct black', async () => {
    const text = record(0, 4000, utf16le('Text'));
    const style = (mask: number, data: Uint8Array) => record(0, 4001, concat(little32(5), little16(0), little32(0), little32(5), little32(mask), data));
    const master = drawing(spContainer(record(0x12, 0xf00a, concat(little32(900), little32(0x800))),
      textbox(header(0), text, style(0x60001, concat(little16(1), little16(36), little32(0xfeffffff))), zeroOriginRuler())));
    const shape = (left: number, active: boolean, black = false, id = 900) => spContainer(
      record(0x12, 0xf00a, concat(little32(left + 1), little32(0xa00 | (active ? 0x20 : 0)))), anchor(576, left, left + 1440, 1440),
      properties([[0x301, id]]), record(15, 0xf011, record(0, 3011, concat(little32(0xffffffff), new Uint8Array([15, 0, 0, 0])))),
      textbox(header(6), text, ...(black ? [style(0x40000, little32(0xfe000000))] : []), zeroOriginRuler()),
    );
    const presentation = await materialize(buildPptFixture(drawing(shape(576, true), shape(2304, true, true), shape(4032, false)), undefined, master));
    const runs = [0, 1, 2].map(index => paragraphs(presentation, index)[0].runs[0]);
    expect(runs[0]).toMatchObject({ text: 'Text', fontSize: 36, bold: true, color: 'FFFFFF' });
    expect(runs[1]).toMatchObject({ text: 'Text', fontSize: 36, bold: true, color: '000000' });
    expect(runs[2]).toMatchObject({ text: 'Text', fontSize: null, bold: null });
    await expect(materialize(buildPptFixture(drawing(shape(576, true, false, 999)), undefined, master))).rejects.toThrow(/unresolved PowerPoint master shape/);
    if (!skia) return;
    // Against an explicit dark background only the inherited white frame shows white ink.
    presentation.slides[0].background = { fillType: 'solid', color: '10263F' };
    const pixels = await render(presentation);
    expect(count(pixels, 96, 96, 240, 144, (r, g, b) => r > 240 && g > 240 && b > 240)).toBeGreaterThan(100);
    expect(count(pixels, 384, 96, 240, 144, (r, g, b) => r > 240 && g > 240 && b > 240)).toBe(0);
  });

  it('inherits master character bullets and honours a direct no-bullet override', async () => {
    const text = record(0, 4000, utf16le('Item'));
    const masterStyle = record(0, 4001, concat(little32(5), little16(0), little32(0x5ed), little16(13),
      little16(0x25a0), little16(100), little32(0xfe0000ff), little16(144), little16(0), little32(5), little32(0x20000), little16(20)));
    const master = drawing(spContainer(record(0x12, 0xf00a, concat(little32(900), little32(0x800))), textbox(header(1), text, masterStyle)));
    const disabled = record(0, 4001, concat(little32(5), little16(0), little32(1), little16(0), little32(5), little32(0)));
    const shape = (left: number, noBullet: boolean) => spContainer(record(0x12, 0xf00a, concat(little32(left + 1), little32(0xa20))),
      anchor(576, left, left + 1440, 1440), properties([[0x301, 900]]), textbox(header(1), text, ...(noBullet ? [disabled] : [])));
    const presentation = await materialize(buildPptFixture(drawing(shape(576, false), shape(2304, true)), undefined, master));
    expect(paragraphs(presentation, 0)[0]).toMatchObject({ marL: 228600, indent: -228600, bullet: { type: 'char', char: '■', color: 'FF0000', sizePct: 100 } });
    expect(paragraphs(presentation, 1)[0].bullet).toMatchObject({ type: 'none' });
    if (!skia) return;
    const pixels = await render(presentation);
    const red = (r: number, g: number, b: number) => r > 200 && g < 80 && b < 80;
    expect(count(pixels, 96, 96, 240, 144, red)).toBeGreaterThan(50);
    expect(count(pixels, 384, 96, 240, 144, red)).toBe(0);
  });

  it('keeps distinct default tab intervals through the model and painting', async () => {
    const shape = (id: number, top: number, tab: number) => spContainer(shapeAtom(202, id, 0x200), anchor(top, 576, 4032, top + 576), textbox(
      header(1), record(0, 4000, utf16le('A\tB')),
      record(0, 4001, concat(little32(4), little16(0), little32(0x8000), little16(tab),
        little32(1), little32(0x20000), little16(24), little32(3), little32(0x60000), little16(24), little32(0xfe0000ff))),
      zeroOriginRuler(),
    ));
    const presentation = await materialize(buildPptFixture(drawing(shape(1, 576, 288), shape(2, 1440, 576))));
    expect([0, 1].map(index => paragraphs(presentation, index)[0].defTabSz)).toEqual([457200, 914400]);
    if (!skia) return;
    const pixels = await render(presentation);
    const redStart = (top: number) => {
      let left = 960;
      for (let y = top; y < top + 96; y++) for (let x = 0; x < 960; x++) {
        const at = (y * 960 + x) * 4;
        if (pixels[at] > 200 && pixels[at + 1] < 80 && pixels[at + 2] < 80 && pixels[at + 3] > 0) left = Math.min(left, x);
      }
      expect(left).toBeLessThan(960);
      return left;
    };
    // The slide is 10 inches wide: a half-inch tab difference is 48 px.
    expect(redStart(240) - redStart(96)).toBe(48);
  });

  it('places each text frame at its own anchor', async () => {
    const shape = (x: number, y: number, width: number, text: string) => spContainer(shapeAtom(202, 42, 0x200),
      // MS-PPT ClientAnchor is top, left, right, bottom in 1/576 inch.
      anchor(y, x, x + width, y + 576), textbox(record(0, 4000, utf16le(text)), zeroOriginRuler()));
    const presentation = await materialize(buildPptFixture(drawing(shape(576, 576, 1728, 'First'), shape(2880, 1728, 1152, 'Second'))));
    expect(presentation.slides[0].elements).toMatchObject([
      { x: 914400, y: 914400, width: 2743200, height: 914400 },
      { x: 4572000, y: 2743200, width: 1828800, height: 914400 },
    ]);
    if (!skia) return;
    const pixels = await render(presentation);
    const ink = (r: number, _g: number, _b: number, a: number) => r < 128 && a > 0;
    expect(count(pixels, 96, 96, 288, 96, ink)).toBeGreaterThan(0);
    expect(count(pixels, 480, 288, 192, 96, ink)).toBeGreaterThan(0);
    expect(count(pixels, 48, 48, 400, 40, ink)).toBe(0);
  });
});

async function render(presentation: Presentation): Promise<Uint8ClampedArray> {
  const canvas = new skia!.Canvas(960, 720);
  await renderSlideNode(canvas as never, presentation, 0, { width: 960, dpr: 1 });
  return canvas.getContext('2d').getImageData(0, 0, 960, 720).data;
}

function count(pixels: Uint8ClampedArray, left: number, top: number, width: number, height: number,
  match: (r: number, g: number, b: number, a: number) => boolean): number {
  let total = 0;
  for (let y = top; y < top + height; y++) for (let x = left; x < left + width; x++) {
    const i = (y * 960 + x) * 4;
    if (match(pixels[i], pixels[i + 1], pixels[i + 2], pixels[i + 3])) total++;
  }
  return total;
}
