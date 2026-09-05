import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildDocFixture, concat, little16, little32 } from '../../legacy-converter/src/test-fixtures.ts';
import { openDocxDocument } from './docx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import type { NodeCanvasFactory } from './render.ts';

const skia = await loadSkiaForTests();
function record(kind: number, options: number, bytes: Uint8Array): Uint8Array {
  return concat(little16(options), little16(kind), little32(bytes.length), bytes);
}
function inlinePicture(png: Uint8Array, cropLeft: number): Uint8Array {
  const properties = concat(little16(0x102), little32(cropLeft), little16(0xc104), little32(0xffffffff));
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(1), little32(0x800))),
    record(0xf00b, (2 << 4) | 3, properties),
  ));
  const blip = record(0xf01e, 0x6e0 << 4, concat(new Uint8Array(17), png));
  const header = new Uint8Array(68);
  const view = new DataView(header.buffer);
  view.setUint32(0, header.length + shape.length + blip.length, true);
  for (const [offset, value] of [[4, 68], [6, 100], [28, 1440], [30, 720], [32, 1000], [34, 1000]]) view.setUint16(offset, value, true);
  return concat(header, shape, blip);
}

it.skipIf(!skia).each([0, 16384])('renders inline binary Word PNG and fixed-point crop through the ordinary DOCX renderer (%i)', async crop => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = { createCanvas: (w, h) => new Canvas(w, h), loadImage: b => loadImage(Buffer.from(new Uint8Array(b))) };
  const source = new Canvas(40, 20);
  const ctx = source.getContext('2d');
  ctx.fillStyle = '#ff0000'; ctx.fillRect(0, 0, 20, 20);
  ctx.fillStyle = '#0000ff'; ctx.fillRect(20, 0, 20, 20);
  const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
  const bytes = buildDocFixture({ text: '\u0001', data: inlinePicture(await source.toBuffer('png'), crop), characterProperties: concat(little16(0x0855), new Uint8Array([1]), little16(0x6a03), little32(0)) });
  const session = await openDocxDocument(bytes, { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
  try {
    expect(session.pageCount).toBe(1);
    const canvas = await session.renderPage(0, { dpr: 1 }) as InstanceType<typeof Canvas>;
    const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    let red = 0, blue = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      if (pixels[i] > 240 && pixels[i + 1] < 15 && pixels[i + 2] < 15) red++;
      if (pixels[i] < 15 && pixels[i + 1] < 15 && pixels[i + 2] > 240) blue++;
    }
    // Display extent is one inch by half an inch at 96 dpi. Allow only
    // antialiased edge pixels, not omission or rescaling of the whole picture.
    expect(red + blue).toBeGreaterThan(4300);
    expect(red + blue).toBeLessThanOrEqual(96 * 48);
    if (crop === 0) expect(red / blue).toBeCloseTo(1, 1);
    else expect(red / blue).toBeCloseTo(0.5, 1);
  } finally { await session.close(); }
});
