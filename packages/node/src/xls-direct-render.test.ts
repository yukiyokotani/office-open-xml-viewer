import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { buildXlsFixture, concat, little16 } from '../../legacy-converter/src/test-fixtures.ts';
import { createLegacyOfficeWasmConverter } from '../../legacy-converter/src/index.ts';
import { materializeXlsxWorkbook } from './xlsx.ts';
import { loadSkiaForTests } from './test-imports.ts';
import { renderWorksheetViewport } from '../../xlsx/src/render-orchestrator.ts';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';

const skia = await loadSkiaForTests();

it.skipIf(!skia)('paints direct XLS cells identically to the byte-conversion route on the same renderer', async () => {
  const { Canvas, loadImage } = skia as typeof import('skia-canvas');
  const bytes = buildXlsFixture({
    sharedString: concat(little16(6), new Uint8Array([0]), new TextEncoder().encode('DIRECT')),
  });
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const directory = (view.getUint32(48, true) + 1) * 2 ** view.getUint16(30, true);
  // The authored fixture's sole Workbook stream must belong to the root.
  view.setUint32(directory + 76, 1, true);
  const converter = createLegacyOfficeWasmConverter({
    wasm: await readFile(new URL('../../legacy-converter/src/wasm/legacy_office_converter_bg.wasm', import.meta.url)),
  });
  const direct = await materializeXlsxWorkbook(bytes, {
    legacyConversion: { xls: { source: {
      protocol: 'ooxml-legacy-xls-source/v1', builtin: 'xls',
      wasmUrl: new URL('../../legacy-converter/src/wasm-direct-xls/legacy_office_converter_bg.wasm', import.meta.url).href,
    } } },
  });
  const previous = await materializeXlsxWorkbook(bytes, { legacyConversion: { xls: { converter } } });
  const factory = {
    createCanvas: (width: number, height: number) => new Canvas(width, height),
    loadImage: (buffer: ArrayBuffer) => loadImage(Buffer.from(new Uint8Array(buffer))),
  };
  const restoreImage = installImageBitmapShim(factory);
  const restoreCanvas = installOffscreenCanvasShim(factory);
  try {
    const paint = async (model: typeof direct) => {
      const canvas = new Canvas(480, 180);
      await renderWorksheetViewport({ ws: model.worksheets[0], styles: model.workbookIndex.styles },
        canvas as unknown as HTMLCanvasElement, { row: 1, col: 1, rows: 5, cols: 5 },
        { width: 480, height: 180, dpr: 1 });
      return canvas.getContext('2d').getImageData(0, 0, 480, 180).data;
    };
    const actual = await paint(direct);
    expect(actual).toEqual(await paint(previous));
    expect(direct.worksheets[0].rows.flatMap(row => row.cells).map(cell => cell.value)).toContainEqual({ type: 'text', text: 'DIRECT' });
    expect(actual.some((value, index) => index % 4 !== 3 && value < 128)).toBe(true);
  } finally { restoreCanvas(); restoreImage(); }
});
