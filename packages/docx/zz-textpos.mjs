// usage: node textpos.mjs <port> <candRoot> <outJson> <rel...>
// Records every fillText word position (in points) per page of a DOCX rendered by the docx VRT fixture.
import { chromium } from '@playwright/test';
import { writeFileSync } from 'node:fs';
const [port, root, outJson, ...rels] = process.argv.slice(2);
const browser = await chromium.launch({ channel: 'chrome', args: ['--disable-gpu'] });
const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, deviceScaleFactor: 1 });
await ctx.addInitScript(() => {
  window.__tp = [];
  const P = CanvasRenderingContext2D.prototype; const orig = P.fillText;
  P.fillText = function (text, x, y, maxWidth) {
    try {
      if (this.canvas && this.canvas.id === 'canvas' && /\S/.test(text)) {
        const m = this.getTransform();
        const toks = []; const re = /\S+/g; let r;
        while ((r = re.exec(text))) {
          const pre = this.measureText(text.slice(0, r.index)).width;
          const w = this.measureText(r[0]).width;
          const lx = x + pre, rx = x + pre + w;
          toks.push({ t: r[0], x0: m.a * lx + m.c * y + m.e, x1: m.a * rx + m.c * y + m.e, y: m.b * x + m.d * y + m.f, scale: m.a });
        }
        window.__tp.push(...toks);
      }
    } catch {}
    return orig.call(this, text, x, y, maxWidth);
  };
});
const result = {};
for (const rel of rels) {
  const page = await ctx.newPage(); page.setDefaultTimeout(240000);
  await page.goto(`http://127.0.0.1:${port}/tests/visual/fixture.html?file=${encodeURIComponent(rel)}&page=0&width=794`);
  await page.waitForFunction(() => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error');
  result[rel] = await page.evaluate(async () => {
    const ds = document.body.dataset; if (ds.status === 'error') return { error: ds.errorMessage };
    const canvas = document.getElementById('canvas'); const n = Number(ds.pageCount); const pages = [];
    for (let i = 0; i < n; i++) {
      window.__tp = []; await window.renderDocxVrtPage(i);
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      pages.push({ widthPx: canvas.width, toks: window.__tp });
    }
    return { pages };
  });
  await page.close();
}
writeFileSync(outJson, JSON.stringify(result));
await browser.close();
