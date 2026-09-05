import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { buildDocFixture } from '../../legacy-converter/src/test-fixtures.ts';
import { openDocxDocument } from './docx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import type { NodeCanvasFactory } from './render.ts';

const skia = await loadSkiaForTests();
const converter = createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });

it.skipIf(!skia).each(['footnotes', 'endnotes'] as const)('renders recovered binary %s through the ordinary OOXML note flow', async kind => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory: NodeCanvasFactory = { createCanvas: (w, h) => new Canvas(w, h), loadImage: b => loadImage(Buffer.from(new Uint8Array(b))) };
  const bytes = buildDocFixture({ text: 'Body\u0002\r', [kind]: [{ cp: 4, text: '\u0002Recovered note\r' }],
    characterProperties: new Uint8Array([0x55, 0x08, 1]), defaultTabTwips: 720 });
  const session = await openDocxDocument(bytes, { factory, currentDate: 0, legacyConversion: { doc: { converter } } });
  try {
    const runs: { text: string; y: number }[] = [];
    for (let page = 0; page < session.pageCount; page++) await session.renderPage(page, { dpr: 1, onTextRun: r => runs.push({ text: r.text, y: r.y + page * 2000 }) });
    expect(runs.map(r => r.text).join('')).toContain('Recovered note');
    const body = runs.find(r => r.text.includes('Body'));
    const note = runs.find(r => r.text.includes('Recovered'));
    expect(note).toBeDefined(); expect(body).toBeDefined();
    expect((note as { y: number }).y).toBeGreaterThan((body as { y: number }).y);
    expect(runs.filter(r => r.text === '1').length).toBeGreaterThanOrEqual(2);
  } finally { await session.close(); }
});
