// A direct XLS workbook painted by the ordinary XLSX canvas renderer in Node.
import { expect, it } from 'vitest';
import { testXlsSource } from '../test-sources.js';
import {
  installImageBitmapShim, installOffscreenCanvasShim, materializeXlsxWorkbook,
  renderWorksheetViewport, skia, skiaFactory,
} from './node-facade.js';
import { buildXlsRichFixture } from './xls-rich-fixture.js';

it.skipIf(!skia)('paints direct XLS cell text and its rich run color through the XLSX canvas path', async () => {
  const factory = skiaFactory();
  const parsed = await materializeXlsxWorkbook(buildXlsRichFixture(), { modelSources: [testXlsSource()], factory });
  const canvas = factory.createCanvas(700, 200);
  const restoreImage = installImageBitmapShim(factory);
  const restoreOffscreen = installOffscreenCanvasShim(factory);
  try {
    await renderWorksheetViewport({ ws: parsed.worksheets[0]!, styles: parsed.workbookIndex.styles },
      canvas as unknown as HTMLCanvasElement, { row: 1, col: 1, rows: 6, cols: 10 }, { width: 700, height: 200, dpr: 1 });
    const context = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const pixels = context.getImageData(0, 0, 700, 200).data;
    let red = 0;
    let dark = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      const [r, g, b] = [pixels[i]!, pixels[i + 1]!, pixels[i + 2]!];
      if (r > 220 && g < 30 && b < 30) red++;
      else if (r < 100 && g < 100 && b < 100) dark++;
    }
    // "RED " paints in its run color; the default-colored runs and the
    // number paint dark text.
    expect(red).toBeGreaterThan(0);
    expect(dark).toBeGreaterThan(0);
  } finally { restoreOffscreen(); restoreImage(); }
});
