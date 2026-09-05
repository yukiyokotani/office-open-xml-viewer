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
const array = (size: number, items: Uint8Array[]) => concat(little16(items.length), little16(items.length), little16(size), ...items);
const atom = (parent: number) => record(2, 1007, concat(new Uint8Array(12), little32(parent), new Uint8Array(4), little16(1), little16(0)));
const drawing = (shapes: Uint8Array) => record(15, 1036, record(15, 0xf002, shapes));
const anchor = record(0, 0xf010, concat(little32(576), little32(576), little32(1728), little32(1728)));
const vertices = array(8, [[10, 120], [10, 20], [110, 20], [110, 120], [10, 120]].map(([x, y]) => concat(little32(x), little32(y))));
const segments = array(2, [0x4000, 0x2001, 1, 0x6001, 0x8000].map(little16));
const props = (values: [number, number][], extra: Uint8Array = new Uint8Array()) => record((values.length << 4) | 3, 0xf00b, concat(...values.map(([id, v]) => concat(little16(id), little32(v))), extra));
const shape = (id: number, kind: number, properties: Uint8Array, flags = 0xa00) => record(15, 0xf004, concat(record((kind << 4) | 2, 0xf00a, concat(little32(id), little32(flags))), anchor, properties));

async function converted(inherited = false) {
  const geometry: [number, number][] = [[0x140, 10], [0x141, 20], [0x142, 110], [0x143, 120], [0x144, 4],
    [0xc145, vertices.length], [0xc146, segments.length]];
  const master = concat(atom(0), drawing(concat(shape(900, 1, props([[0x181, 255], [0x1ff, 0x00080000]])),
    ...(inherited ? [shape(901, 0, props([...geometry, [0x3bf, 0x00020002]], concat(vertices, segments)))] : []))));
  const front = shape(10, 0, props([...(inherited ? [[0x301, 901] as [number, number]] : geometry),
    [0x181, 0xff0000], [0x1ff, 0x00080000]], inherited ? new Uint8Array() : concat(vertices, segments)), inherited ? 0xa20 : 0xa00);
  const wasm = await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url));
  return materializePptxPresentation(buildPptFixture(concat(atom(100), drawing(front)), undefined, master),
    { legacyConversion: { ppt: { converter: createLegacyOfficeWasmConverter({ wasm }) } } });
}

it.each([false, true])('converts explicit or linked master cubic paths into ordinary OOXML custom geometry (inherited=%s)', async inherited => {
  const model = await converted(inherited);
  expect(model.slides[0].elements).toHaveLength(2);
  expect(model.slides[0].elements[1]).toMatchObject({ type: 'shape', geometry: 'custGeom', fill: { color: '0000FF' } });
});

it.skipIf(!skia)('occludes master objects only inside the restored foreground curve', async () => {
  const { Canvas } = skia as typeof import('skia-canvas');
  const canvas = new Canvas(960, 720);
  await renderSlideNode(canvas, await converted(), 0, { width: 960, dpr: 1 });
  const pixel = (x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data);
  expect(pixel(192, 240)).toEqual([0, 0, 255, 255]);
  expect(pixel(110, 110)).toEqual([255, 0, 0, 255]);
  expect(pixel(320, 240)).toEqual([255, 255, 255, 255]);
});
