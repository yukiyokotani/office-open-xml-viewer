import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildDocFixture, concat, little16, little32 } from '../../legacy-converter/src/test-fixtures.ts';
import { openDocxDocument } from './docx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import type { NodeCanvasFactory } from './render.ts';

const skia = await loadSkiaForTests();
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
const format = (value: number) => concat(little16(0x300e), new Uint8Array([value]));
const restart = (value: number) => concat(little16(0x3011), new Uint8Array([1]), little16(0x7044), little32(value));

it.skipIf(!skia).each([
  { properties: [concat(format(2), restart(5)), new Uint8Array(), concat(format(3), restart(2))], expected: ['v', '6', 'B'] },
  { properties: [restart(0), concat(little16(0x501c), little16(99)), format(2)], expected: ['0', '1', 'ii'] },
  { properties: [restart(65536), new Uint8Array(), restart(2147483646)], expected: ['65536', '65537', '2147483646'] },
  { properties: [format(0xff), format(0x16), format(0x17)], expected: ['', '02', '3'] },
])('renders section-local binary page numbering: $expected', async ({ properties, expected }) => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = { createCanvas: (w, h) => new Canvas(w, h), loadImage: b => loadImage(Buffer.from(new Uint8Array(b))) };
  const footer = 'Page \u0013PAGE\u001499\u0015 of \u0013NUMPAGES\u001499\u0015\r';
  const source = buildDocFixture({ text: 'One\fTwo\fThree\r', sectionEnds: [4, 8, 14],
    sectionProperties: properties, defaultTabTwips: 720,
    headers: ['', '', '', footer, '', '', ...Array<string>(12).fill('')],
  });
  const session = await openDocxDocument(source, { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
  try {
    expect(session.pageCount).toBe(3);
    for (let page = 0; page < 3; page++) {
      const runs: { text: string; y: number }[] = [];
      await session.renderPage(page, { dpr: 1, onTextRun: run => runs.push(run) });
      expect(runs.filter(r => r.y > 900).map(r => r.text).join('')).toBe(`Page ${expected[page]} of 3`);
    }
  } finally { await session.close(); }
});

it.skipIf(!skia)('propagates a bounded page-number expansion failure and permits a later open', async () => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = { createCanvas: (w, h) => new Canvas(w, h), loadImage: b => loadImage(Buffer.from(new Uint8Array(b))) };
  const source = (value: number) => buildDocFixture({ text: 'Body\r', defaultTabTwips: 720,
    sectionProperties: concat(format(3), restart(value)),
    headers: ['', '', '', '\u0013PAGE\u001499\u0015\r', '', ''],
  });
  const render = async (value: number) => {
    const session = await openDocxDocument(source(value), { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
    try { await session.renderPage(0, { dpr: 1 }); }
    finally { await session.close(); }
  };
  await expect(render(2147483646)).rejects.toThrow(/number-format output budget/i);
  await expect(render(1)).resolves.toBeUndefined();
});
