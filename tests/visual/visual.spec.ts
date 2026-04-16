import { test } from '@playwright/test';
import { readFileSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

// Number of slides to test (must have corresponding reference images)
const SLIDE_COUNT = 5;

// Per-pixel color tolerance for pixelmatch (0 = exact, 1 = fully lenient)
const PIXEL_THRESHOLD = 0.15;

// Set to a number (e.g. 5) to fail the test when diff exceeds that percentage.
// Set to null to always pass (report-only mode).
const FAIL_ABOVE_PCT: number | null = null;

test.describe('visual regression', () => {
  for (let i = 0; i < SLIDE_COUNT; i++) {
    const slideNum = i + 1;

    test(`slide ${slideNum}`, async ({ page }) => {
      // ── Load the fixture and wait for rendering to complete ──────────────
      await page.goto(`/tests/visual/fixture.html?slide=${i}`);

      await page.waitForFunction(
        () =>
          document.body.dataset.status === 'ready' ||
          document.body.dataset.status === 'error',
        { timeout: 30_000 }
      );

      const status = await page.evaluate(() => document.body.dataset.status);
      if (status === 'error') {
        const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
        throw new Error(`Fixture reported error on slide ${slideNum}: ${msg}`);
      }

      // Give the browser one extra frame to flush composite layers
      await page.waitForTimeout(200);

      // ── Capture the canvas via toDataURL (more reliable than screenshot API) ─
      const dataUrl = await page.evaluate(() => {
        const canvas = document.querySelector('canvas') as HTMLCanvasElement;
        return canvas ? canvas.toDataURL('image/png') : null;
      });
      if (!dataUrl) throw new Error(`No canvas element found on slide ${slideNum}`);
      // dataUrl is "data:image/png;base64,<base64data>"
      const base64 = dataUrl.split(',')[1];
      const actualBuf = Buffer.from(base64, 'base64');
      writeFileSync(`tests/visual/screenshots/slide-${slideNum}.png`, actualBuf);

      // ── Load reference ───────────────────────────────────────────────────
      const refBuf = readFileSync(`tests/visual/references/slide-${slideNum}.png`);
      const refPng    = PNG.sync.read(refBuf);
      const actualPng = PNG.sync.read(actualBuf);

      const { width: refW, height: refH } = refPng;

      // Dimensions must match — if not, report clearly
      if (actualPng.width !== refW || actualPng.height !== refH) {
        console.error(
          `  slide ${slideNum}: size mismatch ` +
          `actual=${actualPng.width}×${actualPng.height} ` +
          `ref=${refW}×${refH}`
        );
      }

      const w = Math.min(actualPng.width, refW);
      const h = Math.min(actualPng.height, refH);

      // ── Pixel comparison ─────────────────────────────────────────────────
      const diff = new PNG({ width: w, height: h });
      const diffPixels = pixelmatch(
        refPng.data,
        actualPng.data,
        diff.data,
        w, h,
        { threshold: PIXEL_THRESHOLD, includeAA: true }
      );
      writeFileSync(`tests/visual/diffs/slide-${slideNum}.png`, PNG.sync.write(diff));

      const totalPx  = w * h;
      const diffPct  = (diffPixels / totalPx) * 100;
      const matchPct = 100 - diffPct;

      // ── Report ───────────────────────────────────────────────────────────
      console.log(
        `  slide ${slideNum}: ` +
        `match=${matchPct.toFixed(1)}%  ` +
        `diff=${diffPct.toFixed(1)}%  ` +
        `(${diffPixels.toLocaleString()} / ${totalPx.toLocaleString()} px)`
      );

      // ── Optional hard failure ─────────────────────────────────────────────
      if (FAIL_ABOVE_PCT !== null && diffPct > FAIL_ABOVE_PCT) {
        throw new Error(
          `Slide ${slideNum} pixel diff ${diffPct.toFixed(1)}% exceeds ` +
          `threshold ${FAIL_ABOVE_PCT}%`
        );
      }
    });
  }
});
