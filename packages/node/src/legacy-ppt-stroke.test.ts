/// <reference types="vite/client" />
import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildPptFixture, concat, little16, little32 } from '../../legacy-converter/src/test-fixtures.ts';
import { materializePptxPresentation } from './pptx.ts';
import { renderSlideNode } from './render.ts';
import { loadSkiaForTests } from './test-imports.ts';

const skia = await loadSkiaForTests();
const record = (options: number, kind: number, payload: Uint8Array) => concat(little16(options), little16(kind), little32(payload.length), payload);
async function converted(properties: [number, number][]) {
  const values: [number, number][] = [[0x1c0, 0xff0000], [0x1cb, 114300], ...properties];
  const shape = record(15, 0xf004, concat(
    record((20 << 4) | 2, 0xf00a, concat(little32(10), little32(0xa00))),
    record(0, 0xf010, concat(little32(576), little32(576), little32(1728), little32(576))),
    record((values.length << 4) | 3, 0xf00b, concat(...values.map(([id, value]) => concat(little16(id), little32(value))))),
  ));
  const bytes = buildPptFixture(record(15, 1036, record(15, 0xf002, shape)));
  const wasm = await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));
  return materializePptxPresentation(bytes, { legacyConversion: { ppt: { converter: createLegacyOfficeWasmConverter({ wasm }) } } });
}

it('passes line-end sizes and cap/join styles through the ordinary PPTX parser', async () => {
  const model = await converted([[0x1d0, 1], [0x1d1, 5], [0x1d2, 0], [0x1d3, 2], [0x1d4, 2], [0x1d5, 0], [0x1d6, 1], [0x1d7, 1], [0x1cc, 0x18000]]);
  expect(model.slides[0].elements[0]).toMatchObject({ type: 'shape', stroke: {
    color: '0000FF', width: 114300, lineCap: 'square', lineJoin: 'miter', miterLimit: 1.5,
    headEnd: { type: 'triangle', w: 'sm', len: 'lg' }, tailEnd: { type: 'arrow', w: 'lg', len: 'sm' },
  } });
});

it.skipIf(!skia)('paints a round cap beyond the endpoint without changing the line anchor', async () => {
  const { Canvas } = skia as typeof import('skia-canvas');
  const pixels: number[][] = [];
  for (const cap of [0, 2]) {
    const model = await converted([[0x1d7, cap]]);
    const canvas = new Canvas(960, 720);
    await renderSlideNode(canvas, model, 0, { width: 960, dpr: 1 });
    pixels.push(Array.from(canvas.getContext('2d').getImageData(92, 96, 1, 1).data));
  }
  expect(pixels).toEqual([[0, 0, 255, 255], [255, 255, 255, 255]]);
});
