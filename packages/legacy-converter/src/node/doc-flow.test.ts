// Binary Word section text flow (MS-DOC 2.6.4 sprmSTextFlow) through the
// ordinary DOCX Canvas renderer.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { openDocxDocument, skia, skiaFactory } from './node-facade.js';

type Pixels = { width: number; height: number; getContext(kind: '2d'): { getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray } } };

it.skipIf(!skia).each([
  '天地玄黄宇宙洪荒\r日月盈昃辰宿列張\r',
  'ABCDEFGH\rIJKLMNOP\r',
  '天地ABC123。\r玄黄DEF456、\r',
])('renders horizontal and vertical section writing direction: %j', async (text) => {
  for (const flow of [0, 1]) {
    const bytes = buildDocFixture({ text, sectionProperties: concat(little16(0x5033), little16(flow)) });
    const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
    try {
      expect(session.pageCount).toBe(1);
      const canvas = await session.renderPage(0, { dpr: 1 }) as unknown as Pixels;
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
