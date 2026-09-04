import { DocxDocument } from '@silurus/ooxml/docx';
import { XlsxSheetViewer, XlsxWorkbook } from '@silurus/ooxml/xlsx';
import { PptxPresentation } from '@silurus/ooxml/pptx';
import { math } from '@silurus/ooxml/math';
import { threeD } from '@silurus/ooxml/three-d';
import { regionMap } from '@silurus/ooxml/region-map';
import { chartEx } from '@silurus/ooxml/chart-ex';
import { createLegacyOfficeWasmWorkerConverter } from '@silurus/ooxml/legacy-conversion';

const renderers = { math, threeD, regionMap, chartEx };
const paint = (id, bitmap) => {
  const canvas = document.getElementById(id);
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  canvas.getContext('2d').drawImage(bitmap, 0, 0);
  bitmap.close();
};
const bytes = async (url) => {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${url}: ${response.status}`);
  return response.arrayBuffer();
};

try {
  document.body.dataset.stage = 'legacy-converter-worker';
  const legacyConverter = createLegacyOfficeWasmWorkerConverter();
  try {
    await legacyConverter.convert({
      bytes: new Uint8Array([1, 2, 3]),
      from: 'doc',
      to: 'docx',
      maxOutputBytes: 1024 * 1024,
      signal: new AbortController().signal,
    });
    throw new Error('legacy converter unexpectedly accepted malformed CFB');
  } catch (error) {
    if (error?.reason !== 'unsupported-input') throw error;
    document.body.dataset.legacyConverter = 'ready';
  }

  const docx = await DocxDocument.load(
    await bytes('/packages/docx/public/demo/sample-1.docx'),
    { mode: 'worker', ...renderers },
  );
  if (docx.mode !== 'worker') throw new Error(`DOCX effective mode: ${docx.mode}`);
  paint('docx', await docx.renderPageToBitmap(0, { width: 360, dpr: 1 }));
  docx.destroy();

  const equation = await DocxDocument.load(
    await bytes('/consumer/equation.docx'),
    { mode: 'worker', ...renderers },
  );
  paint('math', await equation.renderPageToBitmap(0, { width: 360, dpr: 1 }));
  equation.destroy();

  const xlsx = await XlsxWorkbook.load(
    await bytes('/packages/xlsx/public/demo/sample-1.xlsx'),
    { mode: 'worker', ...renderers },
  );
  paint('xlsx', await xlsx.renderViewportToBitmap(
    0,
    { row: 1, col: 1, rows: 20, cols: 10 },
    { width: 360, height: 240, dpr: 1 },
  ));
  xlsx.destroy();

  const pptx = await PptxPresentation.load(
    await bytes('/packages/pptx/public/demo/sample-1.pptx'),
    { mode: 'worker', ...renderers },
  );
  paint('pptx', await pptx.renderSlideToBitmap(0, { width: 360, dpr: 1 }));
  pptx.destroy();

  const textPptx = await PptxPresentation.load(
    await bytes('/consumer/text.pptx'),
    { mode: 'worker', useGoogleFonts: false },
  );
  paint('pptx-text', await textPptx.renderSlideToBitmap(0, { width: 640, dpr: 1 }));
  window.pptxTextRuns = await textPptx.collectSlideRuns(0, 640);
  textPptx.destroy();

  const borderedXlsx = await XlsxWorkbook.load(
    await bytes('/consumer/bordered.xlsx'),
    { mode: 'worker', useGoogleFonts: false },
  );
  paint('xlsx-bordered', await borderedXlsx.renderViewportToBitmap(
    0,
    { row: 0, col: 0, rows: 8, cols: 4 },
    { width: 320, height: 200, dpr: 1 },
  ));
  borderedXlsx.destroy();

  const csvSource = new TextEncoder().encode(
    'id,description,value\r\n00123,"production, worker",=1+1\r\n',
  ).buffer;
  const csvMain = new XlsxSheetViewer(document.getElementById('xlsx-csv-main'), {
    showScrollbars: false,
  });
  document.body.dataset.stage = 'csv-main';
  await csvMain.load(csvSource.slice(0), { format: 'csv' });
  const csvWorker = new XlsxSheetViewer(document.getElementById('xlsx-csv-worker'), {
    mode: 'worker',
    showScrollbars: false,
    onReady: () => { document.body.dataset.stage = 'csv-worker-render'; },
  });
  document.body.dataset.stage = 'csv-worker';
  await csvWorker.load(csvSource.slice(0), { format: 'csv' });
  window.csvViewers = [csvMain, csvWorker];

  const chartExDocx = await DocxDocument.load(
    await bytes('/consumer/chart-ex.docx'),
    { mode: 'worker', ...renderers },
  );
  paint('docx-chart-ex', await chartExDocx.renderPageToBitmap(0, { width: 640, dpr: 1 }));
  chartExDocx.destroy();

  const chartExXlsx = await XlsxWorkbook.load(
    await bytes('/consumer/chart-ex.xlsx'),
    { mode: 'worker', ...renderers },
  );
  paint('xlsx-chart-ex', await chartExXlsx.renderViewportToBitmap(
    0,
    { row: 0, col: 0, rows: 20, cols: 10 },
    { width: 640, height: 360, dpr: 1 },
  ));
  chartExXlsx.destroy();

  const chartExPptx = await PptxPresentation.load(
    await bytes('/consumer/chart-ex.pptx'),
    { mode: 'worker', ...renderers },
  );
  paint('pptx-chart-ex', await chartExPptx.renderSlideToBitmap(0, { width: 640, dpr: 1 }));
  chartExPptx.destroy();

  document.body.dataset.status = 'ready';
} catch (error) {
  document.body.dataset.status = 'error';
  const detail = error instanceof Error ? (error.stack ?? error.message) : String(error);
  document.body.dataset.errorMessage = `${document.body.dataset.stage ?? 'unknown'}: ${detail}`;
  console.error(error);
}
