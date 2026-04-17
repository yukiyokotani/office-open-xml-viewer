import { test } from '@playwright/test';
import init, { parse_pptx } from '../../src/wasm/pptx_parser.js';
import { readFileSync } from 'fs';

test('check parsed data', async ({ page }) => {
  const messages: string[] = [];
  page.on('console', msg => messages.push(msg.text()));
  
  await page.goto('http://localhost:5173/tests/visual/fixture.html?pptx=sample-1&slide=0');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 30_000 }
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
    console.log('ERROR:', msg);
    return;
  }
  await page.waitForTimeout(200);
  // Check pixel at (10,10) - should be inside left sidebar (light blue)
  const pixel = await page.evaluate(() => {
    const canvas = document.querySelector('canvas') as HTMLCanvasElement;
    if (!canvas) return 'no canvas';
    const ctx = canvas.getContext('2d')!;
    const d = ctx.getImageData(10, 10, 1, 1).data;
    return `r=${d[0]} g=${d[1]} b=${d[2]} a=${d[3]}`;
  });
  console.log('OUTPUT: Pixel at (10,10):', pixel);
  // Also check center pixel
  const centerPixel = await page.evaluate(() => {
    const canvas = document.querySelector('canvas') as HTMLCanvasElement;
    const ctx = canvas.getContext('2d')!;
    const d = ctx.getImageData(640, 360, 1, 1).data;
    return `r=${d[0]} g=${d[1]} b=${d[2]} a=${d[3]}`;
  });
  console.log('OUTPUT: Pixel at center:', centerPixel);
  const canvasSize = await page.evaluate(() => {
    const canvas = document.querySelector('canvas') as HTMLCanvasElement;
    return `${canvas.width}x${canvas.height} style=${canvas.style.width}x${canvas.style.height}`;
  });
  console.log('OUTPUT: Canvas size:', canvasSize);
});
