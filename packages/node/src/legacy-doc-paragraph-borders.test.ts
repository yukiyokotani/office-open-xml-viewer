import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildDocFixture, concat, little16 } from '../../legacy-converter/src/test-fixtures.ts';
import { openDocxDocument } from './docx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import type { NodeCanvasFactory } from './render.ts';

const skia = await loadSkiaForTests();
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });

it.skipIf(!skia).each([false, true])('renders binary paragraph rules in body and headers, old=%s', async old => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = { createCanvas: (w, h) => new Canvas(w, h), loadImage: b => loadImage(Buffer.from(new Uint8Array(b))) };
  // Red two-point bottom border; identical adjacent paragraphs form one group.
  const border = old ? concat(little16(0x6426), new Uint8Array([16, 1, 6, 2]))
    : concat(little16(0xc650), new Uint8Array([8, 255, 0, 0, 0, 16, 1, 2, 0]));
  const source = buildDocFixture({ text: 'One\rTwo\r', paragraphProperties: border,
    headers: ['', 'Running header\r', '', '', '', ''], defaultTabTwips: 720,
  });
  const session = await openDocxDocument(source, { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
  try {
    expect(session.pageCount).toBe(1);
    const canvas = await session.renderPage(0, { dpr: 1 }) as InstanceType<typeof Canvas>;
    const { data } = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
    const rows: number[] = [];
    for (let y = 0; y < canvas.height; y++) {
      let red = 0;
      for (let x = 0; x < canvas.width; x++) {
        const i = (y * canvas.width + x) * 4;
        if (data[i] > 200 && data[i + 1] < 80 && data[i + 2] < 80 && data[i + 3] > 200) red++;
      }
      if (red > canvas.width / 2) rows.push(y);
    }
    const groups = rows.filter((y, i) => i === 0 || y !== rows[i - 1] + 1);
    // One header rule and one rule below the pair, not one per body paragraph.
    expect(groups).toHaveLength(2);
    expect(groups[0]).toBeLessThan(96);
    expect(groups[1]).toBeGreaterThan(96);
  } finally { await session.close(); }
});
