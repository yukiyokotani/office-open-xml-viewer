import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildDocFixture, concat, little16 } from '../../legacy-converter/src/test-fixtures.ts';
import { openDocxDocument } from './docx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import type { NodeCanvasFactory } from './render.ts';

const skia = await loadSkiaForTests();
it.skipIf(!skia).each([
  '天地玄黄宇宙洪荒\r日月盈昃辰宿列張\r',
  'ABCDEFGH\rIJKLMNOP\r',
  '天地ABC123。\r玄黄DEF456、\r',
])('renders binary Word section writing direction through the ordinary DOCX Canvas path: %j', async (text) => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = {
    createCanvas: (w, h) => new Canvas(w, h),
    loadImage: b => loadImage(Buffer.from(new Uint8Array(b))),
  };
  const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
  for (const flow of [0, 1]) {
    const bytes = buildDocFixture({
      text,
      sectionProperties: concat(little16(0x5033), little16(flow)),
    });
    const session = await openDocxDocument(bytes, { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
    try {
      expect(session.pageCount).toBe(1);
      const canvas = await session.renderPage(0, { dpr: 1 }) as InstanceType<typeof Canvas>;
      expect([canvas.width, canvas.height]).toEqual([816, 1056]);
      const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
      let left = canvas.width, right = -1, top = canvas.height, bottom = -1;
      for (let y = 0; y < canvas.height; y++) {
        for (let x = 0; x < canvas.width; x++) {
          const i = (y * canvas.width + x) * 4;
          if (pixels[i + 3] > 200 && pixels[i] < 100 && pixels[i + 1] < 100 && pixels[i + 2] < 100) {
            left = Math.min(left, x); right = Math.max(right, x);
            top = Math.min(top, y); bottom = Math.max(bottom, y);
          }
        }
      }
      expect(right).toBeGreaterThan(left);
      expect(bottom).toBeGreaterThan(top);
      if (flow === 1) expect(bottom - top).toBeGreaterThan(right - left);
      else expect(right - left).toBeGreaterThan(bottom - top);
    } finally { await session.close(); }
  }
});
