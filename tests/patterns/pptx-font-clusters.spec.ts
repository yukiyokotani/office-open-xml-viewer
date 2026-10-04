import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

const entry = fileURLToPath(new URL('../../packages/pptx/src/renderer.ts', import.meta.url));

test('PPTX preserves attached Myanmar marks across slot and authored run seams', async ({ page }) => {
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'pptxRenderer' }, platform: 'browser' });
  await page.addScriptTag({ content: bundle.output[0].code });
  const font = (await readFile(new URL('./fixtures/myanmar/paint.woff2', import.meta.url))).toString('base64');
  const results = await page.evaluate(async (encoded) => {
    const bytes = Uint8Array.from(atob(encoded), ch => ch.charCodeAt(0));
    for (const name of ['Cluster CS', 'Cluster EA']) {
      document.fonts.add(await new FontFace(name, bytes).load());
    }
    const renderer = (globalThis as typeof globalThis & {
      pptxRenderer: typeof import('../../packages/pptx/src/renderer.js');
    }).pptxRenderer;
    const results = [];
    // An ordinary extender and the measured extension marks must all retain
    // the base context. Separate-grapheme spacing marks are outside this gate.
    for (const mark of ['\u109d', '\ua9e5', '\uaa7c']) {
      for (const seam of [false, true]) {
        const text = `\u1000${mark}`;
        const run = (value: string) => ({ type: 'text', text: value, lang: 'my-MM',
          fontFamily: 'sans-serif', fontFamilyCs: 'Cluster CS', fontFamilyEa: 'Cluster EA',
          fontSize: 32, bold: false, italic: false, underline: false, strikethrough: false,
          color: '000000' });
        const body = { verticalAnchor: 't', defaultFontSize: 32,
          defaultBold: false, defaultItalic: false, lIns: 0, rIns: 0, tIns: 0, bIns: 0,
          wrap: 'square', vert: 'horz', autoFit: 'none', paragraphs: [{
            alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null,
            spaceLine: null, lvl: 0, bullet: { type: 'none' }, defFontSize: null,
            defColor: null, defBold: null, defItalic: null, defFontFamily: null,
            tabStops: [], eaLnBrk: true, runs: seam ? [run('\u1000'), run(mark)] : [run(text)],
          }] };
        const canvas = document.createElement('canvas');
        canvas.width = 160; canvas.height = 100;
        const ctx = canvas.getContext('2d')!;
        ctx.font = '32px "Cluster CS"';
        const attachedAdvance = ctx.measureText(text).width;
        const separateAdvance = ctx.measureText('\u1000').width + ctx.measureText(mark).width;
        const width = (attachedAdvance + separateAdvance) / 2;
        const calls: { text: string; font: string; x: number; y: number }[] = [];
        const fill = ctx.fillText.bind(ctx);
        ctx.fillText = (value, x, y) => { calls.push({ text: value, font: ctx.font, x, y }); fill(value, x, y); };
        renderer.renderTextBody(ctx, body as never, 0, 0, width, 100, 1 / 12700);
        const reference = document.createElement('canvas');
        reference.width = canvas.width; reference.height = canvas.height;
        const oracle = reference.getContext('2d')!;
        const first = calls[0];
        oracle.font = first.font;
        oracle.fillText(text, first.x, first.y);
        const actual = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
        const expected = oracle.getImageData(0, 0, canvas.width, canvas.height).data;
        let different = 0, ink = 0;
        for (let i = 0; i < actual.length; i++) if (actual[i] !== expected[i]) different++;
        for (let i = 3; i < expected.length; i += 4) if (expected[i]) ink++;
        // Independent real shaper control: isolated mark shaping must be
        // distinguishable from attached shaping, so a tofu-only fixture fails.
        const overflow = renderer.naturalWidthExceedsBbox(ctx, body as never, width, 0, 0, 1 / 12700,
          { themeMajorFont: null, themeMinorFont: null, dpr: 1 });
        results.push({ mark, seam, calls: calls.map(call => call.text), different, ink,
          separates: separateAdvance > attachedAdvance, overflow });
      }
    }
    return results;
  }, font);
  for (const result of results) {
    expect(result.ink).toBeGreaterThan(0);
    expect(result.separates).toBe(true);
    expect(result.overflow).toBe(false);
    expect(result.calls).toEqual([`\u1000${result.mark}`]);
    expect(result.different).toBe(0);
  }
});
