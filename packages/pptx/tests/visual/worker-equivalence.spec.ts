import { test, expect } from '@playwright/test';
import { mkdirSync, readdirSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

// Worker mode must produce identical pixels to main mode: same renderer, same
// fonts (the worker loads the same web fonts into its OffscreenCanvas font set),
// different thread. All three slides come out bit-identical (0.000%); the tiny
// uniform tolerance only absorbs rare AA/hinting noise and still fails on a
// dropped element or a font that didn't load in the worker (which would diverge
// by whole tenths of a percent — that was the symptom before the preload fix).
const SLIDES = [0, 1, 2];
const MAX_DIFF_PCT = [0.1, 0.1, 0.1];

test.afterEach(async ({ page }) => {
  await page.evaluate(() => {
    const destroy = (globalThis as unknown as {
      destroyPptxWorkerVrt?: () => void;
    }).destroyPptxWorkerVrt;
    destroy?.();
  }).catch(() => {
    // Navigation or a load failure may already have disposed the page.
  });
});

for (const slide of SLIDES) {
  test(`worker mode matches main mode › demo/sample-1 slide ${slide + 1}`, async ({ page }) => {
    await page.goto(`/tests/visual/worker-fixture.html?pptx=demo/sample-1&slide=${slide}`);
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
    console.log(`  slide ${slide + 1}: worker-vs-main diff ${pct.toFixed(3)}%`);
    expect(pct).toBeLessThanOrEqual(MAX_DIFF_PCT[slide]);
  });
}

const PRIVATE_CORPUS = (
  process.env.VRT_PRIVATE_CORPUS === '1'
  || process.env.VRT_PRIVATE_WORKER_PARITY === '1'
)
  ? readdirSync('public/private/pptx')
      .filter((file) => file.endsWith('.pptx') && !file.startsWith('~$'))
      .map((file) => `pptx/${file}`)
      .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))
  : [];

test.describe('private corpus worker parity', () => {
  for (const file of PRIVATE_CORPUS) {
    test(file, async ({ page }) => {
      test.setTimeout(600_000);
      const stem = file.slice(0, -'.pptx'.length);
      await page.goto(`/tests/visual/worker-fixture.html?pptx=${encodeURIComponent(`private/${stem}`)}&slide=0`);
      await page.waitForFunction(
        () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
        undefined,
        { timeout: 120_000 },
      );
      const status = await page.evaluate(() => document.body.dataset.status);
      if (status === 'error') {
        throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
      }
      const slideCount = Number(await page.evaluate(() => document.body.dataset.slideCount));
      expect(slideCount).toBeGreaterThan(0);
      const differences: string[] = [];
      for (let slideIndex = 0; slideIndex < slideCount; slideIndex++) {
        if (slideIndex > 0) {
          await page.evaluate(async (index) => {
            const render = (globalThis as unknown as {
              renderPptxWorkerVrtSlide(slideIndex: number): Promise<void>;
            }).renderPptxWorkerVrtSlide;
            await render(index);
          }, slideIndex);
        }
        const [mainUrl, workerUrl] = await page.evaluate(() => [
          (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
          (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
        ]);
        const main = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
        const worker = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
        if (main.width !== worker.width || main.height !== worker.height) {
          differences.push(`${stem} slide-${slideIndex + 1}: dimensions differ`);
          continue;
        }
        const diff = pixelmatch(
          main.data,
          worker.data,
          undefined,
          main.width,
          main.height,
          { threshold: 0.1 },
        );
        const pct = (diff / (main.width * main.height)) * 100;
        // Main draws SVGs as vectors while the worker receives a bounded
        // display-sized bitmap, so edge antialiasing is not byte-identical.
        // Keep the same narrow tolerance as the public worker-parity suite;
        // dropped content changes materially more than this raster edge noise.
        if (pct > 0.1) {
          const output = `tests/visual/screenshots/worker-parity/${stem}`;
          mkdirSync(output, { recursive: true });
          writeFileSync(`${output}/slide-${slideIndex + 1}-main.png`, PNG.sync.write(main));
          writeFileSync(`${output}/slide-${slideIndex + 1}-worker.png`, PNG.sync.write(worker));
          differences.push(`${stem} slide-${slideIndex + 1}: worker/main diff ${pct.toFixed(3)}%`);
        }
      }
      expect(differences, differences.join('\n')).toEqual([]);
    });
  }
});

// IX-nav (M2): the internal-navigation map is derived in BOTH modes — main from
// the parsed slides, worker from the `partNames` that ride through the meta. A
// serialization drop or an order mismatch would make an internal slide-jump land
// on the wrong slide in worker mode only. Assert the resolved-index arrays for
// every slide part name are identical across the two modes (and non-degenerate:
// at least one part name resolved, so we are actually comparing a populated map).
test('worker mode matches main mode › getSlideIndexByPartName map (IX-nav)', async ({ page }) => {
  await page.goto('/tests/visual/worker-fixture.html?pptx=demo/sample-1&slide=0');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }
  const [main, worker] = await page.evaluate(() => [
    document.body.dataset.navMain ?? '[]',
    document.body.dataset.navWorker ?? '[]',
  ]);
  const mainIdx = JSON.parse(main) as number[];
  const workerIdx = JSON.parse(worker) as number[];
  console.log(`  nav map main=${main} worker=${worker}`);
  // A populated map (at least one slide part resolved, not all -1).
  expect(mainIdx.some((i) => i >= 0)).toBe(true);
  // Byte-for-byte identical resolution in both modes.
  expect(workerIdx).toEqual(mainIdx);
});
