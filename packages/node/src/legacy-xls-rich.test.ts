import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { buildXlsFixture, concat, little16 } from '../../legacy-converter/src/test-fixtures.ts';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { materializeXlsxWorkbook, openXlsxWorkbook } from './xlsx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import { renderWorksheetViewport } from '../../xlsx/src/render-orchestrator.ts';

const skia = await loadSkiaForTests();

function fixture(): Uint8Array {
  const record = (kind: number, data: Uint8Array) => concat(little16(kind), little16(data.length), data);
  const fonts: Uint8Array[] = [];
  for (let i = 0; i < 5; i++) {
    const font = new Uint8Array(21);
    const view = new DataView(font.buffer);
    view.setUint16(0, i === 4 ? 480 : 220, true);
    view.setUint16(4, i === 4 ? 10 : 0x7fff, true);
    view.setUint16(6, i === 1 ? 700 : 400, true);
    font[14] = 5; font.set(new TextEncoder().encode('Arial'), 16);
    fonts.push(record(0x31, font));
  }
  const xf = new Uint8Array(20); xf[0] = 1;
  const text = new TextEncoder().encode('base RED normal');
  return buildXlsFixture({
    sharedString: concat(little16(text.length), new Uint8Array([8]), little16(3), text,
      little16(5), little16(5), little16(9), little16(0), little16(15), little16(65535)),
    styleRecords: concat(...fonts, record(0xe0, xf)),
  });
}

async function converter() {
  return createLegacyOfficeWasmConverter({ wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)) });
}

it('preserves separate BIFF run fonts through the opt-in converter and ordinary XLSX parser', async () => {
  const workbook = await openXlsxWorkbook(fixture(), { legacyConversion: { xls: { converter: await converter() } } });
  try {
    const cells = [];
    for await (const chunk of workbook.worksheetRows(0)) {
      if (chunk.kind === 'rows') for (const row of chunk.rows) cells.push(...row.cells);
    }
    const value = cells.find(c => c.row === 2 && c.col === 2)?.value;
    expect(value).toMatchObject({ type: 'text', text: 'base RED normal', runs: [
      { text: 'base ' },
      { text: 'RED ', font: { name: 'Arial', size: 24, color: '#FF0000', bold: false } },
      { text: 'normal', font: { name: 'Arial', size: 11, bold: false } },
    ] });
  } finally { await workbook.close(); }
});

it.skipIf(!skia)('paints the preserved rich run color through the ordinary XLSX Canvas path', async () => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const factory = {
    createCanvas: (w: number, h: number) => new Canvas(w, h),
    loadImage: (bytes: ArrayBuffer) => loadImage(Buffer.from(new Uint8Array(bytes))),
  };
  const parsed = await materializeXlsxWorkbook(fixture(), { legacyConversion: { xls: { converter: await converter() } } });
  const canvas = new Canvas(700, 200);
  const restoreImage = installImageBitmapShim(factory);
  const restoreOffscreen = installOffscreenCanvasShim(factory);
  try {
    await renderWorksheetViewport({ ws: parsed.worksheets[0], styles: parsed.workbookIndex.styles },
      canvas as unknown as HTMLCanvasElement, { row: 1, col: 1, rows: 6, cols: 10 }, { width: 700, height: 200, dpr: 1 });
    const pixels = canvas.getContext('2d').getImageData(0, 0, 700, 200).data;
    let redPixels = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      if (pixels[i] > 220 && pixels[i + 1] < 30 && pixels[i + 2] < 30) redPixels++;
    }
    expect(redPixels).toBeGreaterThan(0);
  } finally { restoreOffscreen(); restoreImage(); }
});
