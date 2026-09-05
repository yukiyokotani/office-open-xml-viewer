import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildDocFixture, concat, little16 } from '../../legacy-converter/src/test-fixtures.ts';
import { openDocxDocument } from './docx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import type { NodeCanvasFactory } from './render.ts';

const skia = await loadSkiaForTests();
it.skipIf(!skia).each([false, true])('renders binary Word header variants and page fields with lock=%s', async lockedHeaderFields => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = { createCanvas: (w, h) => new Canvas(w, h), loadImage: b => loadImage(Buffer.from(new Uint8Array(b))) };
  const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
  const footer = 'Page \u0013PAGE\u001499\u0015 of \u0013NUMPAGES\u001499\u0015\r';
  const source = buildDocFixture({ text: 'One\fTwo\fThree\r',
    headers: ['EVEN\r', 'ODD\r', footer, footer, 'FIRST\r', footer],
    facingPages: true, defaultTabTwips: 720, lockedHeaderFields,
    sectionProperties: concat(little16(0x300a), new Uint8Array([1])),
  });
  const session = await openDocxDocument(source, { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
  try {
    expect(session.pageCount).toBe(3);
    for (let page = 0; page < 3; page++) {
      const runs: { text: string; y: number }[] = [];
      await session.renderPage(page, { dpr: 1, onTextRun: run => runs.push(run) });
      expect(runs.filter(r => r.y < 96).map(r => r.text).join('')).toBe(['FIRST', 'EVEN', 'ODD'][page]);
      expect(runs.filter(r => r.y > 900).map(r => r.text).join('')).toBe(lockedHeaderFields ? 'Page 99 of 99' : `Page ${page + 1} of 3`);
    }
  } finally { await session.close(); }
});

it.skipIf(!skia)('inherits a header across sections and clears it with an explicit blank', async () => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = { createCanvas: (w, h) => new Canvas(w, h), loadImage: b => loadImage(Buffer.from(new Uint8Array(b))) };
  const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
  const source = buildDocFixture({ text: 'One\fTwo\fThree\r', sectionEnds: [4, 8, 14], defaultTabTwips: 720,
    headers: ['', 'INHERITED\r', '', '', '', '', ...Array<string>(6).fill(''), '', '\r', '', '', '', ''],
  });
  const session = await openDocxDocument(source, { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
  try {
    expect(session.pageCount).toBe(3);
    const headers = [];
    for (let page = 0; page < 3; page++) {
      const runs: { text: string; y: number }[] = [];
      await session.renderPage(page, { dpr: 1, onTextRun: run => runs.push(run) });
      headers.push(runs.filter(r => r.y < 96).map(r => r.text).join(''));
    }
    expect(headers).toEqual(['INHERITED', 'INHERITED', '']);
  } finally { await session.close(); }
});
