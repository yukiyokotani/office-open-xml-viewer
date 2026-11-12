import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { build } from 'vite';
import { existsSync } from 'node:fs';
import { symbolFixtureFontBytes } from '../testing/symbol-font-fixture';

type Api = typeof import('./symbol-face') & typeof import('./symbol-font');

// Real Chromium text measurement: a registered symbol-encoded face, the same
// codes under a Unicode cmap, and an absent family. Host faces (Symbol,
// Wingdings) are reported, and asserted only for their structural invariant.
describe.skipIf(!existsSync(chromium.executablePath()))('symbol face measurement in Chromium', () => {
  let browser: Browser;
  let page: Page;
  beforeAll(async () => {
    const root = new URL('./', import.meta.url).pathname;
    const entry = '/__symbol-face-test.ts';
    const result = await build({
      configFile: false, logLevel: 'error',
      plugins: [{ name: 'symbol-face-test', resolveId: id => id === entry ? id : null,
        load: id => id === entry ? [
          `export * from ${JSON.stringify(`${root}symbol-face.ts`)};`,
          `export * from ${JSON.stringify(`${root}symbol-font.ts`)};`,
        ].join('\n') : null }],
      build: { write: false, minify: false, lib: { entry, formats: ['iife'], name: 'symbolFaceTest' } },
    });
    const outputs = Array.isArray(result) ? result : [result];
    const chunk = outputs.flatMap(output => 'output' in output ? output.output : [])
      .find(output => output.type === 'chunk');
    if (!chunk || chunk.type !== 'chunk') throw new Error('bundle missing');
    browser = await chromium.launch({ channel: 'chromium' });
    page = await browser.newPage();
    await page.setContent('<!doctype html><canvas></canvas>');
    await page.addScriptTag({ content: chunk.code });
    const fonts = {
      symbol: Array.from(symbolFixtureFontBytes({ cmap: 'symbol', first: 0x21, last: 0x7e, name: 'FixtureSym' })),
      unicode: Array.from(symbolFixtureFontBytes({ cmap: 'unicode', first: 0x21, last: 0x7e, name: 'FixtureUni' })),
    };
    await page.evaluate(async (fonts) => {
      for (const [family, bytes] of [['Fixture Symbol Face', fonts.symbol], ['Fixture Unicode Face', fonts.unicode]] as const) {
        const face = new FontFace(family, new Uint8Array(bytes));
        document.fonts.add(face);
        await face.load();
      }
    }, fonts);
  }, 60_000);
  afterAll(async () => { await browser?.close(); });

  it('judges the selected face in the painting context', async () => {
    const result = await page.evaluate(() => {
      const api = (globalThis as unknown as { symbolFaceTest: Api }).symbolFaceTest;
      const ctx = document.querySelector('canvas')!.getContext('2d')!;
      const count = (set: string) => [...set].reduce((n, hex) => n + Number.parseInt(hex, 16).toString(2).replace(/0/g, '').length, 0);
      const describe = (family: string) => {
        const encoding = api.measureSymbolFaceEncoding(ctx, family);
        return encoding.kind === 'known' ? { pua: count(encoding.pua), bare: count(encoding.bare) } : encoding.kind;
      };
      return {
        symbolFace: describe('"Fixture Symbol Face"'),
        unicodeFace: describe('"Fixture Unicode Face"'),
        absent: describe('"Ooxml Absent Symbol Face 7"'),
        hostSymbol: describe('"Symbol"'),
        hostWingdings: describe('"Wingdings"'),
        restoredFont: ctx.font,
      };
    });
    // 0x21..0x7E: every code the face maps is owned in both forms.
    // DEL and SHY cannot be probed and stay in a proven symbol face (PUA only).
    expect(result.symbolFace).toEqual({ pua: 0x7e - 0x21 + 1 + 2, bare: 0x7e - 0x21 + 1 });
    // A Unicode face draws its bare codes as Latin text, never a symbol glyph.
    expect(result.unicodeFace).toEqual({ pua: 0, bare: 0 });
    expect(result.absent).toEqual({ pua: 0, bare: 0 });
    // Host faces: either missing/Unicode-encoded (nothing owned) or a
    // symbol-encoded face owning most of its repertoire.
    for (const host of [result.hostSymbol, result.hostWingdings]) {
      expect(host).not.toBe('unknown');
      const { pua } = host as { pua: number };
      expect(pua === 0 || pua > 180).toBe(true);
    }
    expect(result.restoredFont).toBe('10px sans-serif');
  });

  it('converts only what the selected face lacks, and never to colour', async () => {
    const result = await page.evaluate(() => {
      const api = (globalThis as unknown as { symbolFaceTest: Api }).symbolFaceTest;
      const ctx = document.querySelector('canvas')!.getContext('2d')!;
      const oracle = api.createSymbolFaceOracle(ctx);
      const sel = (family: string) => ({ families: [family], weight: 400, style: 'normal' as const });
      const mono = api.createSymbolMonochromeProbe(ctx, '32px sans-serif');
      return {
        installed: oracle.segments('\uF061a', 'Symbol', sel('Fixture Symbol Face')),
        missing: oracle.segments('\uF061a', 'Symbol', sel('Ooxml Absent Symbol Face 7')),
        missingHeart: oracle.segments('\uF0A9', 'Symbol', sel('Ooxml Absent Symbol Face 7')),
        heartMonochrome: mono('♥\uFE0E'),
        // Wingdings 0x43 is THUMBS UP SIGN (U+1F44D): monochrome only where a
        // text face has it; otherwise the authored code is kept.
        thumbs: oracle.segments('\uF043', 'Wingdings', sel('Ooxml Absent Symbol Face 7')),
        thumbsMonochrome: mono('\u{1F44D}\uFE0E'),
      };
    });
    expect(result.installed).toEqual([{ text: '\uF061a', mapped: false }]);
    // An absent face draws no symbol code: converted text may keep its route.
    expect(result.missing).toEqual([{ text: 'αα', mapped: true, unicodeFace: true }]);
    expect(result.missingHeart).toEqual(result.heartMonochrome
      ? [{ text: '♥\uFE0E', mapped: true, unicodeFace: true }] : [{ text: '\uF0A9', mapped: false }]);
    expect(result.thumbs).toEqual(result.thumbsMonochrome
      ? [{ text: '\u{1F44D}\uFE0E', mapped: true, unicodeFace: true }] : [{ text: '\uF043', mapped: false }]);
  });
});
