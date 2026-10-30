import { expect, test } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';

const packageRoot = fileURLToPath(new URL('../../packages/pptx/', import.meta.url));
const fixturePath = fileURLToPath(new URL('./fixtures/pptx-run-inheritance.pptx', import.meta.url));
let server: ViteDevServer;
let url: string;

test.beforeAll(async () => {
  server = await createServer({
    root: packageRoot,
    configFile: resolve(packageRoot, 'vite.config.ts'),
    server: { host: '127.0.0.1', port: 0, strictPort: false },
  });
  await server.listen();
  const address = server.httpServer?.address();
  if (!address || typeof address === 'string') throw new Error('Vite did not assign a port');
  url = `http://127.0.0.1:${address.port}/tests/visual/run-inheritance-fixture.html`;
});
test.afterAll(async () => { await server?.close(); });

test('partial paragraph defaults reach runs and fields; explicit noFill keeps highlight', async ({ page }) => {
  await page.goto(url);
  await expect(page.locator('body')).toHaveAttribute('data-ready', 'true');
  const bytes = [...await readFile(fixturePath)];
  const results = await page.evaluate(async input => {
    type Rendered = { width: number; height: number; bytes: number[] };
    const render = (window as typeof window & {
      renderRunInheritanceFixture: (bytes: number[], index: number, mode: string) => Promise<Rendered>;
    }).renderRunInheritanceFixture;
    const count = (image: Rendered, region: [number, number, number, number]) => {
      let red = 0, yellow = 0, black = 0;
      for (let y = region[1]; y < region[3]; y++) for (let x = region[0]; x < region[2]; x++) {
        const i = 4 * (y * image.width + x);
        const [r, g, b, a] = image.bytes.slice(i, i + 4);
        if (a < 220) continue;
        if (r > 170 && g < 90 && b > 45 && b < 180) red++;
        if (r > 215 && g > 170 && b < 90) yellow++;
        if (r < 50 && g < 50 && b < 50) black++;
      }
      return { red, yellow, black };
    };
    return Promise.all(['main', 'worker'].map(async mode => {
      const image = await render(input, 0, mode);
      return { mode,
        run: count(image, [40, 30, 950, 180]),
        field: count(image, [40, 190, 950, 360]),
        noFill: count(image, [40, 370, 950, 570]),
      };
    }));
  }, bytes);
  for (const result of results) {
    expect(result.run.red, result.mode).toBeGreaterThan(100);
    expect(result.field.red, result.mode).toBeGreaterThan(100);
    expect(result.noFill.yellow, result.mode).toBeGreaterThan(100);
    expect(result.noFill.red, result.mode).toBe(0);
    expect(result.noFill.black, result.mode).toBe(0);
  }
  expect(results[0].run).toEqual(results[1].run);
  expect(results[0].field).toEqual(results[1].field);
  expect(results[0].noFill).toEqual(results[1].noFill);
});

test('noFill preserves independent outline and patterned underline in flat and warped text', async ({ page }) => {
  await page.goto(url);
  await expect(page.locator('body')).toHaveAttribute('data-ready', 'true');
  const bytes = [...await readFile(fixturePath)];
  const results = await page.evaluate(async input => {
    type Rendered = { width: number; height: number; bytes: number[] };
    const render = (window as typeof window & {
      renderRunInheritanceFixture: (bytes: number[], index: number, mode: string) => Promise<Rendered>;
    }).renderRunInheritanceFixture;
    const count = (image: Rendered, y0: number) => {
      let outline = 0, fg = 0, bg = 0, red = 0;
      for (let y = y0; y < image.height; y++) for (let x = 0; x < image.width; x++) {
        const i = 4 * (y * image.width + x);
        const r = image.bytes[i], g = image.bytes[i + 1], b = image.bytes[i + 2];
        if (r < 45 && g > 80 && g < 185 && b < 110) outline++;
        if (r < 55 && g < 110 && b > 140) fg++;
        if (r > 180 && g > 150 && b < 65) bg++;
        if (r > 170 && g < 90 && b > 45 && b < 180) red++;
      }
      return { outline, fg, bg, red };
    };
    return Promise.all(['main', 'worker'].map(async mode => ({
      mode,
      flat: count(await render(input, 0, mode), 570),
      warp: count(await render(input, 1, mode), 0),
    })));
  }, bytes);
  for (const { mode, flat, warp } of results) {
    for (const region of [flat, warp]) {
      expect(region.outline, mode).toBeGreaterThan(30);
      expect(region.fg, mode).toBeGreaterThan(5);
      expect(region.bg, mode).toBeGreaterThan(5);
      expect(region.red, mode).toBe(0);
    }
  }
  expect(results[0].flat).toEqual(results[1].flat);
  expect(results[0].warp).toEqual(results[1].warp);
});

test('noFill preserves gradient uFill in flat and WordArt text', async ({ page }) => {
  await page.goto(url);
  await expect(page.locator('body')).toHaveAttribute('data-ready', 'true');
  const bytes = [...await readFile(fixturePath)];
  const results = await page.evaluate(async input => {
    type Rendered = { width: number; height: number; bytes: number[] };
    const render = (window as typeof window & {
      renderRunInheritanceFixture: (bytes: number[], index: number, mode: string) => Promise<Rendered>;
    }).renderRunInheritanceFixture;
    const count = (image: Rendered, y0: number, y1: number) => {
      let cyan = 0, magenta = 0, outline = 0;
      for (let y = y0; y < y1; y++) for (let x = 0; x < image.width; x++) {
        const i = 4 * (y * image.width + x);
        const r = image.bytes[i], g = image.bytes[i + 1], b = image.bytes[i + 2];
        if (r < 110 && g > 120 && b > 120) cyan++;
        if (r > 120 && g < 110 && b > 120) magenta++;
        if (r < 50 && g > 80 && g < 185 && b < 110) outline++;
      }
      return { cyan, magenta, outline };
    };
    return Promise.all(['main', 'worker'].map(async mode => {
      const image = await render(input, 1, mode);
      return { mode, flat: count(image, 350, 535), warp: count(image, 535, image.height) };
    }));
  }, bytes);
  for (const result of results) for (const region of [result.flat, result.warp]) {
    expect(region.cyan, result.mode).toBeGreaterThan(5);
    expect(region.magenta, result.mode).toBeGreaterThan(5);
    expect(region.outline, result.mode).toBeGreaterThan(30);
  }
  expect(results[0].flat).toEqual(results[1].flat);
  expect(results[0].warp).toEqual(results[1].warp);
});

test('browser parser carries all inherited attributes through runs, fields, breaks and table cells', async ({ page }) => {
  await page.goto(url);
  await expect(page.locator('body')).toHaveAttribute('data-ready', 'true');
  const bytes = [...await readFile(fixturePath)];
  const parsed = await page.evaluate(async input => {
    const parse = (window as typeof window & {
      parseRunInheritanceFixture: (bytes: number[]) => Promise<any>;
    }).parseRunInheritanceFixture;
    const deck = await parse(input);
    const first = deck.slides[0].elements[0].textBody.paragraphs[0].runs[0];
    const third = deck.slides[2];
    const breaks = third.elements.filter((element: any) => element.type === 'shape')
      .map((element: any) => element.textBody.paragraphs[0]);
    const table = third.elements.find((element: any) => element.type === 'table');
    return { first, breaks, table: table?.rows?.[0]?.cells?.[0]?.textBody?.paragraphs?.[0]?.runs };
  }, bytes);
  expect(parsed.first.characterAttributes).toMatchObject({
    kumimoji: '1', lang: 'ja-JP', altLang: 'en-US', kern: '1200',
    cap: 'none', normalizeH: '1', noProof: '1', smtClean: '0',
    smtId: '17', bmk: 'sample', b: '1', sz: '4800',
  });
  expect(parsed.first.characterChildAttributes.latin).toMatchObject({
    typeface: 'Arial', panose: '020B0604020202020204',
  });
  expect(parsed.first.fontFamilyCs).toBe('Amiri');
  expect(parsed.first.hyperlinkMouseOver).toBe('https://example.test/hover');
  expect(parsed.first.caps).toBe('none');
  expect(parsed.first.glyphFill.fillType).toBe('solid');
  expect(parsed.breaks).toHaveLength(2);
  expect(parsed.breaks[0].runs[1]).toMatchObject({ type: 'break', fontSize: 24 });
  expect(parsed.breaks[1].runs[1]).toMatchObject({ type: 'break', fontSize: 72 });
  expect(parsed.breaks[1].runs[0].fontSize).toBe(24);
  expect(parsed.breaks[1].endRunProperties.fontSize).toBe(96);
  expect(parsed.table).toHaveLength(2);
  for (const run of parsed.table) {
    // A defPPr size (#1620 controls) and fill (#1989 controls) have no effect
    // in PowerPoint; its other character properties still cascade into cell
    // runs and fields.
    expect(run).toMatchObject({ bold: true, fontSize: null, color: null });
    expect(run.characterAttributes.lang).toBe('ja-JP');
    expect(run.fontFamilyCs).toBe('Amiri');
  }
});

test('an a:br after text sizes its line with the preceding run in both browser paint paths', async ({ page }) => {
  await page.goto(url);
  await expect(page.locator('body')).toHaveAttribute('data-ready', 'true');
  const bytes = [...await readFile(fixturePath)];
  const results = await page.evaluate(async input => {
    type Rendered = { width: number; height: number; bytes: number[] };
    const render = (window as typeof window & {
      renderRunInheritanceFixture: (bytes: number[], index: number, mode: string) => Promise<Rendered>;
    }).renderRunInheritanceFixture;
    const bands = (image: Rendered, x0: number, x1: number) => {
      const active: number[] = [];
      for (let y = 25; y < 300; y++) {
        let dark = 0;
        for (let x = x0; x < x1; x++) {
          const i = 4 * (y * image.width + x);
          if (image.bytes[i] < 100 && image.bytes[i + 1] < 100 && image.bytes[i + 2] < 100) dark++;
        }
        if (dark > 10) active.push(y);
      }
      const starts = active.filter((y, i) => i === 0 || y > active[i - 1] + 1);
      return starts;
    };
    return Promise.all(['main', 'worker'].map(async mode => {
      const image = await render(input, 2, mode);
      return { mode, normal: bands(image, 50, 580), large: bands(image, 650, 1230) };
    }));
  }, bytes);
  // PowerPoint sizes a line-break mark that follows text at the size of the
  // run before it, never at its own (#1636 controls; renderer.ts
  // followingMark). The 24 pt and 72 pt breaks after 24 pt text therefore
  // produce the same line boxes, and the 96 pt endParaRPr after the final run
  // does not grow the last line either.
  for (const result of results) {
    expect(result.normal.length, result.mode).toBeGreaterThanOrEqual(2);
    expect(result.large, result.mode).toEqual(result.normal);
  }
  expect(results[0]).toMatchObject({ normal: results[1].normal, large: results[1].large });
});

test('noFill text remains registered for selection in main and worker paths', async ({ page }) => {
  await page.goto(url);
  await expect(page.locator('body')).toHaveAttribute('data-ready', 'true');
  const bytes = [...await readFile(fixturePath)];
  const selections = await page.evaluate(async input => {
    const collect = (window as typeof window & {
      collectRunInheritanceText: (bytes: number[], mode: string) => Promise<Array<{ text: string }>>;
    }).collectRunInheritanceText;
    return Promise.all(['main', 'worker'].map(mode => collect(input, mode)));
  }, bytes);
  for (const runs of selections) {
    expect(runs.map(run => run.text).join(' ')).toContain('INVISIBLE GLYPHS');
    expect(runs.map(run => run.text).join(' ')).toContain('OUTLINE AND UNDERLINE');
  }
  expect(selections[0]).toEqual(selections[1]);
});
