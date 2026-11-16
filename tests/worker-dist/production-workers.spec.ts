import { expect, test } from '@playwright/test';

async function expectWorkerBitmaps(page: import('@playwright/test').Page, url: string) {
  const sourceRequests: string[] = [];
  const renderWorkers = new Set<string>();
  const sourceBodies: string[] = [];
  const ordinaryAssetUrls = new Set<string>();
  page.context().on('request', (request) => {
    const pathname = new URL(request.url()).pathname;
    // The consumer fixture embeds a protocol marker in its test source, so
    // inspect the published dist bodies, where a marker means runtime code.
    if (pathname.startsWith('/dist/') && /\.m?js$/.test(pathname)) {
      ordinaryAssetUrls.add(request.url());
    }
    if (/\/assets\/render-worker-(?!host-)[^/]+\.js$/.test(pathname)) renderWorkers.add(pathname);
    if ((pathname.startsWith('/dist/') || pathname.startsWith('/consumer/'))
      && /model-source|source-worker|worker-source|render-worker-source/.test(pathname)) {
      sourceRequests.push(pathname);
    }
  });
  await page.goto(`${url}${url.includes('?') ? '&' : '?'}pause-sources`);
  await expect.poll(async () => {
    const body = page.locator('body');
    return (await body.getAttribute('data-ordinary-ready')) === 'true'
      || (await body.getAttribute('data-status')) === 'error';
  }, { timeout: 60_000 }).toBe(true);
  // Snapshot ordinary-load requests before resuming source stages. Fetch bodies
  // independently: browser response bodies can disappear when workers terminate,
  // including short-lived workers used during the ordinary-load stage itself.
  for (const assetUrl of [...ordinaryAssetUrls]) {
    const response = await page.request.get(assetUrl);
    try {
      expect(response.ok(), `ordinary asset inspection: ${assetUrl}`).toBe(true);
      const body = await response.text();
      if (body.includes('ooxml-model-source-module/v2')
        || body.includes('model source view default')) sourceBodies.push(new URL(assetUrl).pathname);
    } finally {
      await response.dispose();
    }
  }
  const ordinaryRenderWorkers = [...renderWorkers].sort();
  expect(ordinaryRenderWorkers, 'one render worker per format').toHaveLength(3);
  expect(sourceRequests, 'ordinary OOXML main/worker loads must not fetch source chunks').toEqual([]);
  expect(sourceBodies, 'ordinary OOXML main/worker loads must not fetch source module bodies').toEqual([]);
  // #2248: nor the placeholder code that only a failed part needs.
  expect(
    [...ordinaryAssetUrls].filter((assetUrl) => (
      /\/(error-page|error-placeholder-text|parse-error-placeholder|sheet-error-overlay)-[^/]+\.js$/.test(assetUrl)
    )),
    'healthy OOXML main/worker loads must not fetch part-failure placeholder code',
  ).toEqual([]);
  await expect(page.locator('body')).toHaveAttribute('data-ordinary-ready', 'true');
  await expect(page.locator('body')).toHaveAttribute('data-ordinary-loads',
    'docx-worker,xlsx-worker,pptx-worker,docx-main,xlsx-main,pptx-main');
  await page.evaluate(() => (window as typeof window & { resumeSourceStages?: () => void }).resumeSourceStages?.());
  await expect.poll(
    () => page.locator('body').getAttribute('data-status'),
    { timeout: 60_000 },
  ).not.toBe('loading');
  const status = await page.locator('body').getAttribute('data-status');
  if (status !== 'ready') {
    throw new Error(await page.locator('body').getAttribute('data-error-message') ?? status ?? '');
  }

  for (const id of [
    'docx',
    'math',
    'xlsx',
    'pptx',
    'pptx-text',
    'xlsx-bordered',
    'xlsx-csv-main',
    'xlsx-csv-worker',
    'legacy-doc-main',
    'legacy-doc-worker',
  ]) {
    const ink = await page.locator(`#${id}`).evaluate((canvas: HTMLCanvasElement) => {
      // Worker-backed viewers own a `bitmaprenderer` context, so acquiring a
      // second 2D context on their canvas correctly returns null. Copy the
      // displayed bitmap into a disposable 2D canvas before inspecting pixels.
      const sample = document.createElement('canvas');
      sample.width = canvas.width;
      sample.height = canvas.height;
      const context = sample.getContext('2d');
      if (!context) return 0;
      context.drawImage(canvas, 0, 0);
      const pixels = context.getImageData(0, 0, sample.width, sample.height).data;
      let count = 0;
      for (let offset = 0; offset < pixels.length; offset += 4) {
        if (pixels[offset] < 250 || pixels[offset + 1] < 250 || pixels[offset + 2] < 250) count++;
      }
      return count;
    });
    expect(ink, `${id} worker bitmap should contain ink`).toBeGreaterThan(100);
  }

  await expect(page.locator('body')).toHaveAttribute('data-model-sources', 'ready');
  expect([...renderWorkers].sort(), 'model-source loads reuse the ordinary render-worker assets')
    .toEqual(ordinaryRenderWorkers);

  const pptxTextRuns = await page.evaluate(() => (
    window as typeof window & { pptxTextRuns?: Array<Record<string, unknown>> }
  ).pptxTextRuns);
  expect(pptxTextRuns).toHaveLength(1);
  for (const field of ['inShapeX', 'inShapeY', 'w', 'h', 'fontSize']) {
    expect(
      Number.isFinite(pptxTextRuns?.[0]?.[field]),
      `PPTX worker text run ${field} should be finite`,
    ).toBe(true);
  }

  for (const id of ['docx-chart-ex', 'xlsx-chart-ex', 'pptx-chart-ex']) {
    const coloredInk = await page.locator(`#${id}`).evaluate((canvas: HTMLCanvasElement) => {
      const context = canvas.getContext('2d');
      if (!context) return 0;
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
      let count = 0;
      for (let offset = 0; offset < pixels.length; offset += 4) {
        const red = pixels[offset];
        const green = pixels[offset + 1];
        const blue = pixels[offset + 2];
        if (Math.max(red, green, blue) - Math.min(red, green, blue) > 30) count++;
      }
      return count;
    });
    expect(coloredInk, `${id} worker bitmap should execute the ChartEx painter`)
      .toBeGreaterThan(100);
  }
}

test('published dist starts all three render workers with optional renderers', async ({ page }) => {
  await expectWorkerBitmaps(page, '/');
});

test('Vite consumer bundle preserves all three render workers', async ({ page }) => {
  await expectWorkerBitmaps(page, '/consumer/index.html');
});

// Pixel contracts through the public viewer, not a helper-only Canvas test.
// No corpus snapshots or Office fidelity thresholds are used here.
async function viewerPathShadeSamples(
  page: import('@playwright/test').Page,
  source: string,
  width: number,
  slide: number,
  points: number[][],
) {
  await page.route('**/path-shade-viewer', route => route.fulfill({
    contentType: 'text/html', body: '<!doctype html><body></body>',
  }));
  await page.goto('/path-shade-viewer');
  return page.evaluate(async ({ source, width, slide, points }) => {
    const entry = '/dist/pptx.mjs';
    const { PptxViewer } = await import(entry);
    const response = await fetch(source);
    if (!response.ok) throw new Error(`Control load failed: ${response.status}`);
    const bytes = await response.arrayBuffer();
    const samples = [];
    for (const mode of ['main', 'worker']) {
      const canvas = document.createElement('canvas');
      document.body.append(canvas);
      const viewer = new PptxViewer(canvas, { mode, width, dpr: 1 });
      try {
        await viewer.load(bytes.slice(0));
        if (slide !== 0) await viewer.goToSlide(slide);
        const copy = document.createElement('canvas');
        copy.width = canvas.width;
        copy.height = canvas.height;
        const ctx = copy.getContext('2d') as CanvasRenderingContext2D;
        ctx.drawImage(canvas, 0, 0);
        samples.push(points.map(([x, y]) => [...ctx.getImageData(x, y, 1, 1).data]));
      } finally {
        viewer.destroy();
        canvas.remove();
      }
    }
    return samples;
  }, { source, width, slide, points });
}

test('published PPTX viewer retains rect point paint and shades shape tiles in main and worker modes', async ({ page }) => {
  const samples = await viewerPathShadeSamples(page, '/consumer/path-gradient.pptx', 480, 0,
    [[60, 120], [60, 70], [300, 120], [300, 70]]);
  // Rect point foci retain main's radial field: its diagonal half-box sample
  // is ~180. Shape point foci still use box isolines (~128 at both samples).
  const expected = [128, 180, 128, 128];
  for (const modeSamples of samples) for (const [index, pixel] of modeSamples.entries()) {
    expect(pixel[0]).toBeGreaterThanOrEqual(expected[index] - 4);
    expect(pixel[0]).toBeLessThanOrEqual(expected[index] + 4);
    expect(pixel.slice(1)).toEqual([pixel[0], pixel[0], 255]);
  }
  expect(samples[1]).toEqual(samples[0]);
  expect(Math.abs(samples[0][1][0] - 128)).toBeGreaterThan(30);
  expect(Math.abs(samples[0][2][0] - 90)).toBeGreaterThan(30);
});

test('local shape-path control slide reaches the published PPTX viewer in both modes', async ({ page }) => {
  const { existsSync } = await import('node:fs');
  const file = 'packages/pptx/public/private/pptx/controls-1599/shape-paths.pptx';
  test.skip(!existsSync(file), 'Local-only Office control deck is absent');
  // Slide 4, three-stop RGB rect, centered point focus. Coordinates belong to
  // the control manifest; they sample its axial and diagonal half-box isoline.
  const samples = await viewerPathShadeSamples(page, `/${file}`, 1920, 3,
    [[337, 764], [337, 721]]);
  for (const modeSamples of samples) for (const pixel of modeSamples) {
    expect(pixel[0]).toBeLessThan(12);
    expect(pixel[1]).toBeGreaterThan(243);
    expect(pixel[2]).toBeLessThan(12);
    expect(pixel[3]).toBe(255);
  }
  expect(samples[1]).toEqual(samples[0]);
  // The previous shape brush's half-diagonal radius puts the axial point at
  // s≈.353, between red and green (~75,180,0), not at the green middle stop.
  expect(samples[0][0][1] - 180).toBeGreaterThan(60);
});

// A formatter exception used to abort the public load at the first paint,
// hiding even an unrelated sheet. Exercise real parser/assets in both realms.
test('published XLSX viewer isolates an out-of-range date and retains other sheets', async ({ page }) => {
  const { dateRangeXlsxBytes } = await import('../fixtures/date-range-xlsx.mjs');
  await page.route('**/date-range-viewer', route => route.fulfill({
    contentType: 'text/html', body: '<!doctype html><body></body>',
  }));
  await page.goto('/date-range-viewer');
  const results = await page.evaluate(async bytes => {
    const entry = '/dist/xlsx.mjs';
    const { XlsxViewer, XlsxWorkbook } = await import(entry);
    const results = [];
    for (const mode of ['main', 'worker']) {
      const host = document.createElement('div');
      host.style.cssText = 'width:800px;height:400px';
      document.body.append(host);
      const errors: string[] = [];
      const viewer = new XlsxViewer(host, { mode, onError: (e: Error) => errors.push(e.message) });
      try {
        await viewer.load(new Uint8Array(bytes).buffer);
        await viewer.goToSheet(1);
        const matches = await viewer.findText('Healthy sheet');
        results.push({ mode, errors, matches: matches.length });
      } finally {
        viewer.destroy();
        host.remove();
      }
    }
    const wb = await XlsxWorkbook.load(new Uint8Array(bytes).buffer);
    try {
      const dates = await wb.getWorksheet(0);
      const other = await wb.getWorksheet(1);
      return { results, invalid: wb.cellText(dates, dates.rows[0].cells[0]), healthy: wb.cellText(other, other.rows[0].cells[0]) };
    } finally {
      wb.destroy();
    }
  }, [...dateRangeXlsxBytes()]);
  expect(results).toEqual({
    results: [{ mode: 'main', errors: [], matches: 1 }, { mode: 'worker', errors: [], matches: 1 }],
    invalid: '#', healthy: 'Healthy sheet',
  });
});

// Built-in time IDs need no file-authored numFmt. Check actual parser and
// published main/worker painting plus cellText/find, including subsecond carry.
test('published XLSX resolves built-in time formats without custom numFmts', async ({ page }) => {
  const { builtinTimeXlsxBytes } = await import('../fixtures/builtin-time-xlsx.mjs');
  await page.route('**/builtin-time-viewer', route => route.fulfill({
    contentType: 'text/html', body: '<!doctype html><body></body>',
  }));
  await page.goto('/builtin-time-viewer');
  const results = await page.evaluate(async bytes => {
    const entry = '/dist/xlsx.mjs';
    const { XlsxViewer, XlsxWorkbook } = await import(entry);
    const results = [];
    for (const mode of ['main', 'worker']) {
      const data = new Uint8Array(bytes).buffer;
      const wb = await XlsxWorkbook.load(data.slice(0), { mode });
      let texts;
      try {
        const sheet = await wb.getWorksheet(0);
        texts = sheet.rows.map((row: { cells: unknown[] }) => row.cells.map(cell => wb.cellText(sheet, cell)));
      } finally {
        wb.destroy();
      }
      const host = document.createElement('div');
      host.style.cssText = 'width:800px;height:400px';
      document.body.append(host);
      const viewer = new XlsxViewer(host, { mode });
      try {
        await viewer.load(data);
        const matches = await viewer.findText('1084818:00:00');
        results.push({ mode, texts, matches: matches.length });
      } finally {
        viewer.destroy();
        host.remove();
      }
    }
    return results;
  }, [...builtinTimeXlsxBytes()]);
  const texts = [
    ['18:00', '18:00', '0:01', '0:00'],
    ['18:00:00', '18:00:00', '0:01:02', '0:00:59'],
    ['00:00', '00:00', '01:02', '00:59'],
    ['18:00:00', '1084818:00:00', '0:01:02', '0:00:59'],
    ['00:00.0', '00:00.0', '01:02.3', '01:00.0'],
  ];
  expect(results).toEqual([
    { mode: 'main', texts, matches: 1 },
    { mode: 'worker', texts, matches: 1 },
  ]);
});
