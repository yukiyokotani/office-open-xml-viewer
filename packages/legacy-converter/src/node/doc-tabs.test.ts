// Binary Word tab stops (MS-DOC 2.6.2 sprmPChgTabsPapx, 2.7.2 DopBase.dxaTab)
// through the ordinary DOCX layout and Canvas renderer.
import { expect, it } from 'vitest';
import { buildDocFixture, concat, little16 } from '../test-fixtures.js';
import { testDocSource } from '../test-sources.js';
import { openDocxDocument, skia, skiaFactory } from './node-facade.js';

type Pixels = { width: number; height: number; getContext(kind: '2d'): { getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray } } };

it.skipIf(!skia).each(['custom', 'document-default', 'custom-over-default'] as const)('positions text after %s tabs', async (kind) => {
  const positions: number[] = [];
  for (const tab of [720, 2160]) {
    const bytes = buildDocFixture({
      text: '\tLabel\r',
      ...(kind === 'custom-over-default' ? { defaultTabTwips: 1080 } : {}),
      ...(kind !== 'document-default'
        ? { paragraphProperties: concat(little16(0xc60d), new Uint8Array([5, 0, 1]), little16(tab), new Uint8Array([0])) }
        : { defaultTabTwips: tab }),
    });
    const session = await openDocxDocument(bytes, { factory: skiaFactory(), currentDate: 0, modelSources: [testDocSource()] });
    try {
      expect(session.pageCount).toBe(1);
      const canvas = await session.renderPage(0, { dpr: 1 }) as unknown as Pixels;
      const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
      let minX = canvas.width;
      for (let y = 0; y < canvas.height; y++) {
        for (let x = 0; x < canvas.width; x++) {
          const i = (y * canvas.width + x) * 4;
          if (pixels[i + 3] > 200 && pixels[i] < 100 && pixels[i + 1] < 100 && pixels[i + 2] < 100) minX = Math.min(minX, x);
        }
      }
      expect(minX).toBeLessThan(canvas.width);
      positions.push(minX);
    } finally { await session.close(); }
  }
  // 1,440 twips = 72 pt = 96 Canvas pixels at the default 96-DPI render scale.
  expect(positions[1] - positions[0]).toBe(96);
});
