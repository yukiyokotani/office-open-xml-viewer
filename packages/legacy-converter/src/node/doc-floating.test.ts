// A binary Word floating picture (MS-DOC 2.8.2 PlcfSpa / 2.9.253 Spa and the
// MS-ODRAW drawing group) rendered at its page-relative position.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16, little32 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { openDocxDocument, skia, skiaFactory } from './node-facade.js';

const record = (kind: number, options: number, bytes: Uint8Array): Uint8Array => concat(little16(options), little16(kind), little32(bytes.length), bytes);
type Pixels = { width: number; height: number; getContext(kind: '2d'): { getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray } } };

it.skipIf(!skia).each([1, 2, 3])('renders a floating PNG at its page-relative SPA position (wrap %i)', async wrapping => {
  const { Canvas } = skia as NonNullable<typeof skia>;
  const raster = new Canvas(20, 10); const ctx = raster.getContext('2d'); ctx.fillStyle = '#ff0000'; ctx.fillRect(0, 0, 20, 10);
  const blip = record(0xf01e, 0x6e0 << 4, concat(new Uint8Array(17), await raster.toBuffer('png')));
  const store = record(0xf001, (1 << 4) | 15, blip);
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(1027), little32(0xa00))),
    record(0xf00b, (1 << 4) | 3, concat(little16(0x4104), little32(1))),
    record(0xf010, 0, little32(0)),
  ));
  // Spa: 1-inch left/top, 2x1.5-inch extent, page-relative anchors.
  const floatingAnchors = concat(little32(0), little32(1), little32(1027), little32(1440), little32(1440), little32(2880), little32(2160), little16((1 << 1) | (1 << 3) | (wrapping << 5)), little32(0));
  const drawingGroupData = concat(record(0xf000, 15, store), new Uint8Array([0]), record(0xf002, 15, record(0xf003, 15, shape)));
  const bytes = buildDocFixture({ text: '\u0008', floatingAnchors, drawingGroupData, characterProperties: concat(little16(0x0855), new Uint8Array([1])) });
  const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
  try {
    expect(session.pageCount).toBe(1);
    const canvas = await session.renderPage(0, { dpr: 1 }) as unknown as Pixels;
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
