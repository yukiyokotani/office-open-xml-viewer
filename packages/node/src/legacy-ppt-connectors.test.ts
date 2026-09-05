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
async function converted(anchor: number[], flags = 0xb00) {
  const properties = [[0x181, 255], [0x1c0, 0xff0000], [0x1cb, 38100], [0x1d1, 1]];
  const shape = record(15, 0xf004, concat(
    record((32 << 4) | 2, 0xf00a, concat(little32(10), little32(flags))),
    record(0, 0xf010, concat(...anchor.map(little32))),
    record((properties.length << 4) | 3, 0xf00b, concat(...properties.map(([id, value]) => concat(little16(id), little32(value))))),
  ));
  const bytes = buildPptFixture(record(15, 1036, record(15, 0xf002, shape)));
  const wasm = await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));
  return materializePptxPresentation(bytes, { legacyConversion: { ppt: { converter: createLegacyOfficeWasmConverter({ wasm }) } } });
}

it.each([
  [[576, 576, 1728, 576], 1828800, 0],
  [[576, 576, 576, 1728], 0, 1828800],
] as const)('preserves a straight connector with a zero-sized axis (%j)', async (anchor, width, height) => {
  const model = await converted([...anchor]);
  expect(model.slides[0].elements).toHaveLength(1);
  expect(model.slides[0].elements[0]).toMatchObject({
    type: 'shape', geometry: 'straightConnector1', width, height, fill: { fillType: 'none' },
    stroke: { color: '0000FF', tailEnd: { type: 'triangle' } },
  });
});

it.skipIf(!skia)('paints the connector and its arrow at the authored end, including horizontal reflection', async () => {
  const { Canvas } = skia as typeof import('skia-canvas');
  for (const flipped of [false, true]) {
    const model = await converted([576, 576, 1728, 576], flipped ? 0xb40 : 0xb00);
    const canvas = new Canvas(960, 720);
    await renderSlideNode(canvas, model, 0, { width: 960, dpr: 1 });
    const pixel = (x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data);
    expect(pixel(192, 96)).toEqual([0, 0, 255, 255]);
    expect(pixel(flipped ? 108 : 276, 100)).toEqual([0, 0, 255, 255]);
    expect(pixel(flipped ? 276 : 108, 100)).toEqual([255, 255, 255, 255]);
  }
});
