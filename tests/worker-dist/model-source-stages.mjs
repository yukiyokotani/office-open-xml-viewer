// Model-source stages shared by the published-dist fixture and the Vite
// consumer bundle. They load real OOXML archives through a test-only source
// module by URL (fake-model-source.mjs) and a legacy DOC through the published
// legacy-doc source, in main and worker modes.

const FAKE_MODULE = new URL('/tests/worker-dist/fake-model-source.mjs', location.href).href;

function fakeSource(target, config) {
  return {
    target,
    claim: () => true,
    beginLoad: () => ({
      module: {
        protocol: 'ooxml-model-source-module/v1',
        target,
        moduleUrl: FAKE_MODULE,
        config: { format: target, ...config },
      },
      release() {},
    }),
  };
}

async function docxPage(DocxDocument, source, options) {
  const document = await DocxDocument.load(source.slice(0), options);
  try {
    const canvas = window.document.createElement('canvas');
    await document.renderPage(canvas, 0, { width: 480, dpr: 1 });
    return { image: canvas.toDataURL(), document };
  } catch (error) {
    document.destroy();
    throw error;
  }
}

async function xlsxView(XlsxWorkbook, source, options) {
  const workbook = await XlsxWorkbook.load(source.slice(0), options);
  try {
    const canvas = window.document.createElement('canvas');
    await workbook.renderViewport(canvas, 0, { row: 0, col: 0, rows: 8, cols: 4 }, { width: 320, height: 200, dpr: 1 });
    return canvas.toDataURL();
  } finally {
    workbook.destroy();
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

export async function runModelSourceStages({ DocxDocument, XlsxWorkbook, legacyDocSource, bytes, paintCanvas }) {
  const tracked = await bytes('/consumer/tracked.docx');
  const bordered = await bytes('/consumer/bordered.xlsx');
  for (const mode of ['main', 'worker']) {
    document.body.dataset.stage = `model-source-docx-${mode}`;
    const final = await docxPage(DocxDocument, tracked, { mode });
    final.document.destroy();
    const markup = await docxPage(DocxDocument, tracked, { mode, showTrackedChanges: true });
    markup.document.destroy();
    assert(final.image !== markup.image, 'tracked.docx must render differently in the markup view');
    // The source's view default applies when the caller does not choose.
    const defaulted = await docxPage(DocxDocument, tracked, {
      mode,
      modelSources: [fakeSource('docx', { showTrackedChanges: true, minimal: true })],
    });
    assert(defaulted.image === markup.image, `${mode}: source view default was not applied`);
    // Missing optional capabilities degrade generically.
    const metrics = await defaulted.document.getResourceMetrics();
    assert(metrics && typeof metrics === 'object', `${mode}: metrics without usage must resolve`);
    let markdownError;
    try { await defaulted.document.toMarkdown(); } catch (error) { markdownError = error; }
    assert(/unsupported for this source/.test(String(markdownError?.message)),
      `${mode}: toMarkdown must reject for a source without Markdown`);
    defaulted.document.destroy();
    // An explicit caller choice, including false, wins over the default.
    const explicit = await docxPage(DocxDocument, tracked, {
      mode,
      showTrackedChanges: false,
      modelSources: [fakeSource('docx', { showTrackedChanges: true })],
    });
    explicit.document.destroy();
    assert(explicit.image === final.image, `${mode}: explicit showTrackedChanges:false was not honoured`);

    document.body.dataset.stage = `model-source-xlsx-${mode}`;
    const plain = await xlsxView(XlsxWorkbook, bordered, { mode, useGoogleFonts: false });
    const noRequest = await xlsxView(XlsxWorkbook, bordered, {
      mode,
      useGoogleFonts: false,
      modelSources: [fakeSource('xlsx', { layoutFamily: '' })],
    });
    assert(noRequest === plain, `${mode}: a null host layout request must not change geometry`);
    const wide = await xlsxView(XlsxWorkbook, bordered, {
      mode,
      useGoogleFonts: false,
      modelSources: [fakeSource('xlsx', { layoutFamily: 'Arial', layoutSizePt: 36, minimal: true })],
    });
    assert(wide !== plain, `${mode}: the host-measured Normal font must size the grid`);
    document.body.dataset[`xlsxHostLayout${mode === 'main' ? 'Main' : 'Worker'}`] = wide;
  }
  assert(document.body.dataset.xlsxHostLayoutMain === document.body.dataset.xlsxHostLayoutWorker,
    'main-mode (page) and worker-mode (worker) host layout must measure the same width');
  delete document.body.dataset.xlsxHostLayoutMain;
  delete document.body.dataset.xlsxHostLayoutWorker;

  for (const mode of ['main', 'worker']) {
    document.body.dataset.stage = `legacy-doc-${mode}`;
    const legacy = await DocxDocument.load(await bytes('/consumer/legacy.doc'), {
      mode,
      modelSources: [legacyDocSource()],
    });
    try {
      await legacy.renderPage(paintCanvas(`legacy-doc-${mode}`), 0, { width: 360, dpr: 1 });
    } finally {
      legacy.destroy();
    }
  }
  document.body.dataset.modelSources = 'ready';
}
