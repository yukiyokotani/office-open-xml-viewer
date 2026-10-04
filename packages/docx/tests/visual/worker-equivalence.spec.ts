import { test, expect } from '@playwright/test';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
// @ts-ignore - shared plain-JS test fixture builder.
import { chartExDocxBytes, storedZip } from '../../../../tests/fixtures/chart-ex-packages.mjs';

// Worker mode must produce (near-)identical pixels to main mode: same
// renderer, same pagination, same fonts, different thread. The only expected
// drift is sub-pixel text rasterization between the main-thread <canvas> and
// the worker OffscreenCanvas. Per-page tolerances grant slack only to a page
// that needs it, so the (near-)bit-identical pages keep enough sensitivity
// that even a single dropped text element fails the diff.
const PAGES = [0, 1];
const MAX_DIFF_PCT = [0.2, 0.2];

test('public paragraph geometry preserves empty caret alignment, table sources and page breaks in both modes', async ({ page }) => {
  const continuationText = Array.from({ length: 80 }, (_, index) => `L🙂${String(index).padStart(2, '0')}`);
  const bytes = storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'],
    ['word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
      <w:p w14:paraId="AA000001"><w:pPr><w:jc w:val="right"/></w:pPr></w:p>
      <w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="1440" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p w14:paraId="AA000002"/></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:textDirection w:val="tbRl"/></w:tcPr><w:p w14:paraId="AA000004"/></w:tc></w:tr></w:tbl>
      <w:p w14:paraId="AA000003"><w:pPr><w:pageBreakBefore/></w:pPr></w:p>
      <w:p><w:r><w:t>After empty paragraphs</w:t></w:r></w:p>
      <w:p w14:paraId="AA000005"><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r>${continuationText.map(text => `<w:t>${text}</w:t>`).join("<w:br/>")}</w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
      </w:body></w:document>`],
  ]);
  await page.route('**/paragraph-geometry.docx', route => route.fulfill({ body: bytes,
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
  await page.goto('/tests/visual/paragraph-geometry-fixture.html');
  await page.waitForFunction(() => ['ready', 'error'].includes(document.body.dataset.status ?? ''));
  const result = await page.evaluate(() => ({ ...document.body.dataset }));
  expect(result.status, result.error).toBe('ready');
  expect(result.main).toBe(result.worker);
});

for (const pageIndex of PAGES) {
  test(`worker mode matches main mode › demo/sample-1 page ${pageIndex + 1}`, async ({ page }) => {
    await page.goto(`/tests/visual/worker-fixture.html?docx=demo/sample-1&page=${pageIndex}`);
    // Two full loads (main + worker) plus worker spin-up per test — twice the
    // single-render budget visual.spec.ts uses.
    await page.waitForFunction(
      () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
      { timeout: 60_000 },
    );
    const status = await page.evaluate(() => document.body.dataset.status);
    if (status === 'error') {
      throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
    }

    const [mainUrl, workerUrl] = await page.evaluate(() => [
      (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
      (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    ]);
    const a = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
    const b = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
    // A zero-size canvas means a silently failed render; fail with a readable
    // assertion instead of the NaN the diff percentage would produce.
    expect(a.width).toBeGreaterThan(0);
    expect(a.height).toBeGreaterThan(0);
    expect(b.width).toBe(a.width);
    expect(b.height).toBe(a.height);

    const diff = pixelmatch(a.data, b.data, undefined, a.width, a.height, { threshold: 0.1 });
    const pct = (diff / (a.width * a.height)) * 100;
    console.log(`  page ${pageIndex + 1}: worker-vs-main diff ${pct.toFixed(3)}%`);
    expect(pct).toBeLessThanOrEqual(MAX_DIFF_PCT[pageIndex]);
  });
}

test('worker mode renders non-empty equation pixels through the same math renderer', async ({ page }) => {
  await page.goto('/tests/visual/math-worker-equivalence.html');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }

  const [mainUrl, workerUrl] = await page.evaluate(() => [
    (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
  ]);
  const main = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
  const worker = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
  expect(worker.width).toBe(main.width);
  expect(worker.height).toBe(main.height);
  let inkPixels = 0;
  for (let offset = 0; offset < main.data.length; offset += 4) {
    if (main.data[offset] < 240
      || main.data[offset + 1] < 240
      || main.data[offset + 2] < 240) {
      inkPixels++;
    }
  }
  expect(inkPixels).toBeGreaterThan(500);

  const diff = pixelmatch(
    main.data,
    worker.data,
    undefined,
    main.width,
    main.height,
    { threshold: 0.1 },
  );
  const pct = (diff / (main.width * main.height)) * 100;
  console.log(`  math renderer: worker-vs-main diff ${pct.toFixed(3)}%, ink ${inkPixels}`);
  expect(pct).toBeLessThanOrEqual(0.01);
});

// The optional ChartEx renderer crosses the worker boundary as a descriptor.
// Public demo files carry no ChartEx part, so serve a self-authored waterfall
// and require identical, coloured (not grayscale-placeholder) output.
test('worker mode matches main mode › ChartEx waterfall', async ({ page }) => {
  const bytes = chartExDocxBytes();
  await page.route('**/generated/chart-ex.docx', (route) => route.fulfill({
    status: 200, contentType: 'application/octet-stream', body: Buffer.from(bytes),
  }));
  await page.goto('/tests/visual/worker-fixture.html?docx=generated/chart-ex&page=0');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }
  const [mainUrl, workerUrl] = await page.evaluate(() => [
    (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
  ]);
  const main = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
  const worker = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
  expect(worker.width).toBe(main.width);
  expect(worker.height).toBe(main.height);
  const coloured = (png: PNG): number => {
    let count = 0;
    for (let offset = 0; offset < png.data.length; offset += 4) {
      const r = png.data[offset], g = png.data[offset + 1], b = png.data[offset + 2];
      if (Math.max(r, g, b) - Math.min(r, g, b) > 60) count++;
    }
    return count;
  };
  expect(coloured(main)).toBeGreaterThan(500);
  expect(coloured(worker)).toBeGreaterThan(500);
  const diff = pixelmatch(main.data, worker.data, undefined, main.width, main.height, { threshold: 0.1 });
  expect((diff / (main.width * main.height)) * 100).toBeLessThanOrEqual(0.1);
});
