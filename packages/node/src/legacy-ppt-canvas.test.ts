/// <reference types="vite/client" />
// The source converter's WASM URL import is resolved by Vitest for this test.
import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildPptFixture, concat, little16, little32, utf16le } from '../../legacy-converter/src/test-fixtures.ts';
import { materializePptxPresentation } from './pptx.ts';
import { renderSlideNode } from './render.ts';
import { loadSkiaForTests } from './test-imports.ts';

const skia = await loadSkiaForTests();
const converterWasm = readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));

describe('binary PowerPoint to ordinary Canvas rendering', () => {
  it('retains outline-referenced size and paragraph alignment in the ordinary parser model', async () => {
    const record = (options: number, kind: number, payload: Uint8Array) => concat(little16(options), little16(kind), little32(payload.length), payload);
    const outline = concat(
      record(0, 3999, little32(0)), record(0, 4000, utf16le('Wide')),
      record(0, 4001, concat(
        little32(5), little16(0), little32(0x800), little16(1),
        little32(5), little32(0x20000), little16(48),
      )),
    );
    const drawing = record(15, 1036, record(15, 0xf002, record(15, 0xf004, concat(
      record((202 << 4) | 2, 0xf00a, concat(little32(42), little32(0x200))),
      record(0, 0xf010, concat(little32(576), little32(576), little32(3456), little32(1728))),
      record(15, 0xf00d, record(0, 3998, little32(0))),
    ))));
    const converter = createLegacyOfficeWasmConverter({ wasm: await converterWasm });
    const presentation = await materializePptxPresentation(buildPptFixture(drawing, outline), { legacyConversion: { ppt: { converter } } });
    const element = presentation.slides[0].elements[0];
    expect(element.type).toBe('shape');
    if (element.type !== 'shape') throw new Error('Expected converted text shape');
    expect(element.textBody?.paragraphs[0]).toMatchObject({ alignment: 'ctr', runs: [{ text: 'Wide', fontSize: 48 }] });
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
