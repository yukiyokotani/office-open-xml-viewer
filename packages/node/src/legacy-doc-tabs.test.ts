import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildDocFixture, concat, little16 } from '../../legacy-converter/src/test-fixtures.ts';
import { openDocxDocument } from './docx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import type { NodeCanvasFactory } from './render.ts';

const skia = await loadSkiaForTests();
it.skipIf(!skia)('positions binary Word tabbed text through the ordinary DOCX layout and Canvas renderer', async () => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = { createCanvas: (w,h) => new Canvas(w,h), loadImage: b => loadImage(Buffer.from(new Uint8Array(b))) };
  const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
  const positions: number[] = [];
  for (const tab of [720, 2160]) {
    const bytes = buildDocFixture({ text: '\tLabel\r', paragraphProperties: concat(little16(0xc60d),new Uint8Array([5,0,1]),little16(tab),new Uint8Array([0])) });
    const session = await openDocxDocument(bytes, { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
    try {
      expect(session.pageCount).toBe(1);
      const canvas = await session.renderPage(0, { dpr: 1 }) as InstanceType<typeof Canvas>;
      const pixels = canvas.getContext('2d').getImageData(0,0,canvas.width,canvas.height).data;
      let minX = canvas.width;
      for(let y=0;y<canvas.height;y++)for(let x=0;x<canvas.width;x++){
        const i=(y*canvas.width+x)*4;
        if(pixels[i+3]>200&&pixels[i]<100&&pixels[i+1]<100&&pixels[i+2]<100)minX=Math.min(minX,x);
      }
      expect(minX).toBeLessThan(canvas.width);
      positions.push(minX);
    } finally { await session.close(); }
  }
  // 1,440 twips = 72 pt = 96 Canvas pixels at the default 96-DPI render scale.
  expect(positions[1]-positions[0]).toBe(96);
});
