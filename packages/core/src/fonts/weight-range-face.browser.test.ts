import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser } from '@playwright/test';
import { existsSync, readFileSync } from 'node:fs';

// FontSourceSessionOptions.requestedWeights: a variable file declared at 400
// and at 700 (as Google Fonts serves one) registers as one face with the
// weight range 400-700. Each requested weight must paint and measure exactly
// as with the two faces, in a real engine. Fixture: a synthetic variable
// font (wght 100-900) whose outlines thicken with the weight
// (scripts/generate-fixture-font.py).

const FONT = readFileSync(new URL('../../../../tests/fixtures/fonts/ooxml-fixture-weight-var.ttf', import.meta.url));

async function launch(): Promise<Browser | null> {
  try {
    return existsSync(chromium.executablePath()) ? await chromium.launch() : await chromium.launch({ channel: 'chrome' });
  } catch {
    return null;
  }
}

describe('a weight-range face paints like one face per requested weight', () => {
  let browser: Browser | null;
  beforeAll(async () => { browser = await launch(); });
  afterAll(async () => { await browser?.close(); });

  it('matches pixels and metrics at 400 and 700, also synthesised italic', async (context) => {
    if (!browser) context.skip();
    const page = await browser!.newPage();
    await page.setContent('<!doctype html><body></body>');
    const result = await page.evaluate(async (bytes) => {
      const data = new Uint8Array(bytes).buffer;
      const separate = [new FontFace('Separate', data, { weight: '400' }), new FontFace('Separate', data, { weight: '700' })];
      const ranged = [new FontFace('Ranged', data, { weight: '400 700' })];
      for (const face of [...separate, ...ranged]) { document.fonts.add(face); await face.load(); }
      const paint = (font: string) => {
        const canvas = document.createElement('canvas');
        canvas.width = 640; canvas.height = 80;
        const ctx = canvas.getContext('2d')!;
        ctx.font = font;
        ctx.fillText('Served face line 0123', 4, 56);
        const m = ctx.measureText('Served face line 0123');
        return {
          pixels: Array.from(ctx.getImageData(0, 0, 640, 80).data).join(','),
          metrics: [m.width, m.actualBoundingBoxAscent, m.actualBoundingBoxDescent, m.fontBoundingBoxAscent, m.fontBoundingBoxDescent],
        };
      };
      const out: Record<string, boolean | unknown> = {};
      for (const style of ['', 'italic ']) {
        for (const weight of ['400', '700', 'normal', 'bold']) {
          const a = paint(`${style}${weight} 40px Separate`);
          const b = paint(`${style}${weight} 40px Ranged`);
          const fallback = paint(`${style}${weight} 40px monospace`);
          out[`${style}${weight}`] = { same: a.pixels === b.pixels && JSON.stringify(a.metrics) === JSON.stringify(b.metrics), served: a.pixels !== fallback.pixels };
        }
      }
      // The two weights paint differently (the variation axis is applied).
      out.weightsDiffer = paint('400 40px Ranged').pixels !== paint('700 40px Ranged').pixels;
      // A weight between them does not: why the requested weights must be proven.
      out.betweenDiffers = paint('600 40px Separate').pixels !== paint('600 40px Ranged').pixels;
      return out;
    }, Array.from(FONT));
    for (const [key, value] of Object.entries(result)) {
      if (key === 'weightsDiffer' || key === 'betweenDiffers') expect(value, key).toBe(true);
      else expect(value, key).toEqual({ same: true, served: true });
    }
    await page.close();
  });
});
