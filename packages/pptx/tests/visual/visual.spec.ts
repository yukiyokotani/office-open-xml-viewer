import { test } from '@playwright/test';
import { mkdirSync, readFileSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

// ── Test targets ──────────────────────────────────────────────────────────────
// Add entries here to include additional PPTX files.
// Each entry needs:
//   name       : filename stem (loads /{name}.pptx, reads references/{name}/)
//   slideCount : number of slides to test (must have matching reference images)
const PPTX_FILES: { name: string; slideCount: number }[] = [
  { name: 'private/sample-1', slideCount: 5 },
  { name: 'private/sample-2', slideCount: 17 },
  { name: 'private/sample-3', slideCount: 9 },
  { name: 'private/sample-4', slideCount: 15 },
  { name: 'demo/sample-1', slideCount: 9 },
];

// Per-pixel color tolerance for pixelmatch (0 = exact, 1 = fully lenient)
// 0.20 absorbs font hinting / sub-pixel differences between PowerPoint and Canvas
const PIXEL_THRESHOLD = 0.20;

// Set to a number (e.g. 20) to fail the test when diff exceeds that percentage.
// Set to null to always pass (report-only mode).
const FAIL_ABOVE_PCT: number | null = 20;

// ── Tests ─────────────────────────────────────────────────────────────────────
test.describe('visual regression', () => {
  for (const { name, slideCount } of PPTX_FILES) {
    for (let i = 0; i < slideCount; i++) {
      const slideNum = i + 1;

      test(`${name} › slide ${slideNum}`, async ({ page }) => {
        // ── Load the fixture and wait for rendering to complete ────────────
        await page.goto(`/tests/visual/fixture.html?pptx=${name}&slide=${i}`);

        await page.waitForFunction(
          () =>
            document.body.dataset.status === 'ready' ||
            document.body.dataset.status === 'error',
          { timeout: 30_000 }
        );

        const status = await page.evaluate(() => document.body.dataset.status);
        if (status === 'error') {
          const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
          throw new Error(`Fixture error on ${name} slide ${slideNum}: ${msg}`);
        }

        // Give the browser one extra frame to flush composite layers
        await page.waitForTimeout(200);

        // ── Capture the canvas via toDataURL ──────────────────────────────
        const dataUrl = await page.evaluate(() => {
          const canvas = document.querySelector('canvas') as HTMLCanvasElement;
          return canvas ? canvas.toDataURL('image/png') : null;
        });
        if (!dataUrl) throw new Error(`No canvas on ${name} slide ${slideNum}`);
        const actualBuf = Buffer.from(dataUrl.split(',')[1], 'base64');

        mkdirSync(`tests/visual/screenshots/${name}`, { recursive: true });
        writeFileSync(`tests/visual/screenshots/${name}/slide-${slideNum}.png`, actualBuf);

        // ── Load reference ─────────────────────────────────────────────────
        const refBuf = readFileSync(`tests/visual/references/${name}/slide-${slideNum}.png`);
        const refPng    = PNG.sync.read(refBuf);
        const actualPng = PNG.sync.read(actualBuf);

        const { width: refW, height: refH } = refPng;

        if (actualPng.width !== refW || actualPng.height !== refH) {
          console.error(
            `  ${name} slide ${slideNum}: size mismatch ` +
            `actual=${actualPng.width}×${actualPng.height} ` +
            `ref=${refW}×${refH}`
          );
        }

        const w = Math.min(actualPng.width, refW);
        const h = Math.min(actualPng.height, refH);

        // ── Pixel comparison ───────────────────────────────────────────────
        const diff = new PNG({ width: w, height: h });
        const diffPixels = pixelmatch(
          refPng.data,
          actualPng.data,
          diff.data,
          w, h,
          { threshold: PIXEL_THRESHOLD, includeAA: true }
        );
        mkdirSync(`tests/visual/diffs/${name}`, { recursive: true });
        writeFileSync(`tests/visual/diffs/${name}/slide-${slideNum}.png`, PNG.sync.write(diff));

        const totalPx  = w * h;
        const diffPct  = (diffPixels / totalPx) * 100;
        const matchPct = 100 - diffPct;

        // ── Report ─────────────────────────────────────────────────────────
        console.log(
          `  ${name} slide ${slideNum}: ` +
          `match=${matchPct.toFixed(1)}%  ` +
          `diff=${diffPct.toFixed(1)}%  ` +
          `(${diffPixels.toLocaleString()} / ${totalPx.toLocaleString()} px)`
        );

        // ── Optional hard failure ──────────────────────────────────────────
        if (FAIL_ABOVE_PCT !== null && diffPct > FAIL_ABOVE_PCT) {
          throw new Error(
            `${name} slide ${slideNum} pixel diff ${diffPct.toFixed(1)}% exceeds ` +
            `threshold ${FAIL_ABOVE_PCT}%`
          );
        }
      });
    }
  }
});
