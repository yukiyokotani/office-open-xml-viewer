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
async function converted(dash: number) {
  const properties = [[0x1c0, 0xff0000], [0x1cb, 38100], [0x1ce, dash], [0x1d1, 1]];
  const shape = record(15, 0xf004, concat(
    record((20 << 4) | 2, 0xf00a, concat(little32(10), little32(0xa00))),
    record(0, 0xf010, concat(...[576, 576, 1728, 576].map(little32))),
    record((properties.length << 4) | 3, 0xf00b, concat(...properties.map(([id, value]) => concat(little16(id), little32(value))))),
  ));
  const wasm = await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));
  return materializePptxPresentation(buildPptFixture(record(15, 1036, record(15, 0xf002, shape))),
    { legacyConversion: { ppt: { converter: createLegacyOfficeWasmConverter({ wasm }) } } });
}

it.each([
  'solid', 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot', 'dot',
  'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot',
].map((name, id) => [id, name] as const))('maps binary dash %s to ordinary DrawingML %s', async (id, name) => {
  const model = await converted(id);
  expect(model.slides[0].elements[0]).toMatchObject({ type: 'shape',
    stroke: { color: '0000FF', tailEnd: { type: 'triangle' } },
  });
  const shape = model.slides[0].elements[0];
  if (shape.type !== 'shape') throw new Error('Expected a line shape');
  // The ordinary parser normalizes explicit solid to an absent dashStyle.
  expect(shape.stroke?.dashStyle).toBe(id === 0 ? undefined : name);
});

it.skipIf(!skia)('restores a dashed leader with gaps and its independent solid arrow tip', async () => {
  const { Canvas } = skia as typeof import('skia-canvas');
  const canvas = new Canvas(960, 720);
  await renderSlideNode(canvas, await converted(6), 0, { width: 960, dpr: 1 });
  const ctx = canvas.getContext('2d');
  const row = ctx.getImageData(100, 96, 150, 1).data;
  let blue = 0, white = 0;
  for (let i = 0; i < row.length; i += 4) {
    if (row[i] === 0 && row[i + 1] === 0 && row[i + 2] === 255) blue++;
    if (row[i] === 255 && row[i + 1] === 255 && row[i + 2] === 255) white++;
  }
  expect(blue).toBeGreaterThan(0);
  expect(white).toBeGreaterThan(0);
  expect(Array.from(ctx.getImageData(276, 100, 1, 1).data)).toEqual([0, 0, 255, 255]);
});
