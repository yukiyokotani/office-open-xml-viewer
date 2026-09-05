import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildDocFixture, concat, little16, little32 } from '../../legacy-converter/src/test-fixtures.ts';
import { openDocxDocument } from './docx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import type { NodeCanvasFactory } from './render.ts';

const skia = await loadSkiaForTests();
const record = (kind: number, options: number, bytes: Uint8Array): Uint8Array => concat(little16(options), little16(kind), little32(bytes.length), bytes);

it.skipIf(!skia).each([1, 2, 3])('renders binary Word floating PNG at its page-relative SPA position (wrap %i)', async wrapping => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = { createCanvas: (w, h) => new Canvas(w, h), loadImage: b => loadImage(Buffer.from(new Uint8Array(b))) };
  const raster = new Canvas(20, 10); const ctx = raster.getContext('2d'); ctx.fillStyle = '#ff0000'; ctx.fillRect(0, 0, 20, 10);
  const blip = record(0xf01e, 0x6e0 << 4, concat(new Uint8Array(17), await raster.toBuffer('png')));
  const store = record(0xf001, (1 << 4) | 15, blip);
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(1027), little32(0xa00))),
    record(0xf00b, (1 << 4) | 3, concat(little16(0x4104), little32(1))),
    record(0xf010, 0, little32(0)),
  ));
  const floatingAnchors = concat(little32(0), little32(1), little32(1027), little32(1440), little32(1440), little32(2880), little32(2160), little16((1 << 1) | (1 << 3) | (wrapping << 5)), little32(0));
  const drawingGroupData = concat(record(0xf000, 15, store), new Uint8Array([0]), record(0xf002, 15, record(0xf003, 15, shape)));
  const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
  const bytes = buildDocFixture({ text: '\u0008', floatingAnchors, drawingGroupData, characterProperties: concat(little16(0x0855), new Uint8Array([1])) });
  const session = await openDocxDocument(bytes, { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
  try {
    expect(session.pageCount).toBe(1);
    const canvas = await session.renderPage(0, { dpr: 1 }) as InstanceType<typeof Canvas>;
    const image = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
    let left = canvas.width, right = -1, top = canvas.height, bottom = -1;
    for (let y = 0; y < canvas.height; y++) for (let x = 0; x < canvas.width; x++) {
      const i = (y * canvas.width + x) * 4;
      if (image.data[i] > 240 && image.data[i + 1] < 15 && image.data[i + 2] < 15) {
        left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y);
      }
    }
    expect([left, top, right, bottom]).toEqual([96, 96, 191, 143]);
  } finally { await session.close(); }
});
