#!/usr/bin/env node
// Render public specimens to PNG for review: every page, slide or sheet of
// the named cases (or all cases of a format) through the package's VRT
// fixture with the corpus fonts, plus contact sheets of up to 12 images
// so a whole batch can be checked at a glance.
//
//   node tests/corpus/render.mjs --format docx [--cases docx-024,docx-031] --out DIR [--port N]
//
// Writes <out>/<stem>-<n>.png and <out>/sheet-<k>.png; prints the paths.

import { spawn } from 'node:child_process';
import { mkdirSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from '@playwright/test';

const args = process.argv.slice(2);
const opt = (name, fallback) => {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : fallback;
};
const format = opt('--format');
const ADAPTERS = {
  docx: { ext: 'docx', url: (stem) => `/tests/visual/fixture.html?file=${encodeURIComponent(`corpus/${stem}.docx`)}&width=595&fonts=corpus`, count: 'pageCount', render: 'renderDocxVrtPage' },
  pptx: { ext: 'pptx', url: (stem) => `/tests/visual/fixture.html?pptx=${encodeURIComponent(`corpus/${stem}`)}&fonts=corpus`, count: 'slideCount', render: 'renderPptxVrtSlide' },
  xlsx: { ext: 'xlsx', url: (stem) => `/tests/visual/fixture.html?file=${encodeURIComponent(`corpus/${stem}.xlsx`)}&rows=40&cols=10&width=1280&height=960&fonts=corpus`, count: 'sheetCount', render: 'renderXlsxVrtSheet' },
};
const adapter = ADAPTERS[format];
if (!adapter) throw new Error('--format docx|pptx|xlsx is required');
const out = opt('--out');
if (!out) throw new Error('--out DIR is required');
const port = Number(opt('--port', 5198));
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const pkgDir = join(root, 'packages', format);
const corpusDir = join(pkgDir, 'public', 'corpus');
const all = readdirSync(corpusDir).filter((f) => f.endsWith(`.${adapter.ext}`)).map((f) => f.slice(0, -adapter.ext.length - 1)).sort();
const wanted = opt('--cases')?.split(',').filter(Boolean);
const stems = wanted ? all.filter((s) => wanted.some((w) => s === w || s.startsWith(`${w}-`))) : all;
if (stems.length === 0) throw new Error('no matching cases');
mkdirSync(out, { recursive: true });

const child = spawn('pnpm', ['exec', 'vite', '--host', '127.0.0.1', '--port', String(port), '--strictPort'], { cwd: pkgDir, stdio: 'ignore', detached: true });
const stop = () => { try { process.kill(-child.pid); } catch {} };
const base = `http://127.0.0.1:${port}`;
let up = false;
for (let i = 0; i < 120 && !up; i++) {
  try { up = (await fetch(base + '/tests/visual/fixture.html')).ok; } catch { await new Promise((r) => setTimeout(r, 500)); }
}
if (!up) { stop(); throw new Error('vite did not start'); }

const browser = await chromium.launch();
const images = [];
let failed = 0;
try {
  const page = await browser.newPage();
  for (const stem of stems) {
    try {
      await page.goto(base + adapter.url(stem), { timeout: 60000 });
      await page.waitForFunction(() => ['ready', 'error'].includes(document.body.dataset.status), null, { timeout: 120000 });
      const status = await page.evaluate(() => [document.body.dataset.status, document.body.dataset.errorMessage]);
      if (status[0] === 'error') throw new Error(status[1]);
      const total = Number(await page.evaluate((k) => document.body.dataset[k], adapter.count)) || 1;
      for (let i = 0; i < total; i++) {
        if (i) await page.evaluate(([fn, idx]) => window[fn](idx), [adapter.render, i]);
        const dataUrl = await page.evaluate(() => document.querySelector('canvas')?.toDataURL('image/png'));
        const path = join(out, `${stem}-${i + 1}.png`);
        writeFileSync(path, Buffer.from(dataUrl.split(',')[1], 'base64'));
        images.push({ path, dataUrl });
        console.log(path);
      }
    } catch (e) {
      failed++;
      console.log(`FAILED ${stem}: ${String(e?.message ?? e).split('\n')[0]}`);
    }
  }
  // Contact sheets: 6 × 2 thumbnails, 460 px wide, drawn in the browser.
  for (let k = 0; k < images.length; k += 12) {
    const batch = images.slice(k, k + 12).map((im) => im.dataUrl);
    const png = await page.evaluate(async (urls) => {
      const imgs = await Promise.all(urls.map((u) => new Promise((ok) => { const im = new Image(); im.onload = () => ok(im); im.src = u; })));
      const W = 460;
      const H = Math.max(...imgs.map((im) => Math.round(im.height * W / im.width)));
      const c = document.createElement('canvas');
      c.width = W * 6; c.height = H * 2;
      const g = c.getContext('2d');
      g.fillStyle = '#ddd'; g.fillRect(0, 0, c.width, c.height);
      imgs.forEach((im, i) => g.drawImage(im, (i % 6) * W, Math.floor(i / 6) * H, W - 4, Math.round(im.height * W / im.width) - 4));
      return c.toDataURL('image/png');
    }, batch);
    const path = join(out, `sheet-${k / 12 + 1}.png`);
    writeFileSync(path, Buffer.from(png.split(',')[1], 'base64'));
    console.log(path);
  }
} finally {
  await browser.close();
  stop();
}
if (failed) process.exitCode = 1;
