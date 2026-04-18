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
  // Canvas control is transferred to the worker (OffscreenCanvas), so getContext('2d') is not
  // available on the main-thread element. Use screenshot-based pixel inspection instead.
  const canvasSize = await page.evaluate(() => {
    const canvas = document.querySelector('canvas') as HTMLCanvasElement;
    return `${canvas.width}x${canvas.height} style=${canvas.style.width}x${canvas.style.height}`;
  });
  console.log('OUTPUT: Canvas size:', canvasSize);
});
