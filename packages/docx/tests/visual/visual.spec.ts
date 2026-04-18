import { test } from '@playwright/test';
import { mkdirSync, readFileSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

const DOCX_FILES: { name: string; pageCount: number; width: number }[] = [
  { name: 'sample-1', pageCount: 1, width: 612 },
  { name: 'sample-2', pageCount: 1, width: 595 },
];

const PIXEL_THRESHOLD = 0.20;
const FAIL_ABOVE_PCT: number | null = 20;

test.describe('docx visual regression', () => {
  for (const { name, pageCount, width } of DOCX_FILES) {
    for (let i = 0; i < pageCount; i++) {
      const pageNum = i + 1;

      test(`${name} › page ${pageNum}`, async ({ page }) => {
        await page.goto(
          `/tests/visual/fixture.html?file=${name}.docx&page=${i}&width=${width}`
        );

        await page.waitForFunction(
          () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
          { timeout: 30_000 }
        );

        const status = await page.evaluate(() => document.body.dataset.status);
        if (status === 'error') {
          const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
          throw new Error(`Fixture error on ${name} page ${pageNum}: ${msg}`);
        }

        await page.waitForTimeout(200);

        const dataUrl = await page.evaluate(() => {
          const canvas = document.querySelector('canvas') as HTMLCanvasElement;
          return canvas ? canvas.toDataURL('image/png') : null;
        });
        if (!dataUrl) throw new Error(`No canvas on ${name} page ${pageNum}`);
        const actualBuf = Buffer.from(dataUrl.split(',')[1], 'base64');

        mkdirSync(`tests/visual/screenshots/${name}`, { recursive: true });
        writeFileSync(`tests/visual/screenshots/${name}/page-${pageNum}.png`, actualBuf);

        const refPath = `tests/visual/references/${name}/page-${pageNum}.png`;
        const refBuf = readFileSync(refPath);
        const refPng    = PNG.sync.read(refBuf);
        const actualPng = PNG.sync.read(actualBuf);

        const { width: refW, height: refH } = refPng;

        if (actualPng.width !== refW || actualPng.height !== refH) {
          console.warn(
            `  ${name} page ${pageNum}: size mismatch ` +
            `actual=${actualPng.width}×${actualPng.height} ` +
            `ref=${refW}×${refH}`
          );
        }

        const w = Math.min(actualPng.width, refW);
        const h = Math.min(actualPng.height, refH);

        const diff = new PNG({ width: w, height: h });
        const diffPixels = pixelmatch(
          refPng.data, actualPng.data, diff.data, w, h,
          { threshold: PIXEL_THRESHOLD, includeAA: true }
        );
        mkdirSync(`tests/visual/diffs/${name}`, { recursive: true });
        writeFileSync(`tests/visual/diffs/${name}/page-${pageNum}.png`, PNG.sync.write(diff));

        const totalPx = w * h;
        const diffPct = (diffPixels / totalPx) * 100;
        const matchPct = 100 - diffPct;

        console.log(
          `  ${name} page ${pageNum}: ` +
          `match=${matchPct.toFixed(1)}%  diff=${diffPct.toFixed(1)}%  ` +
          `(${diffPixels.toLocaleString()} / ${totalPx.toLocaleString()} px)`
        );

        if (FAIL_ABOVE_PCT !== null && diffPct > FAIL_ABOVE_PCT) {
          throw new Error(
            `${name} page ${pageNum} pixel diff ${diffPct.toFixed(1)}% exceeds ${FAIL_ABOVE_PCT}%`
          );
        }
      });
    }
  }
});
