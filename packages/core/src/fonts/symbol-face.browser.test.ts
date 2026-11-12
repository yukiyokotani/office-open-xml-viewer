import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { build } from 'vite';
import { existsSync } from 'node:fs';
import { symbolFixtureFontBytes } from '../testing/symbol-font-fixture';

type Api = typeof import('./symbol-face') & typeof import('./symbol-font')
  & Pick<typeof import('./route-sentinels'), 'loadNativeRouteSentinels'>;

// Real Chromium: the oracle asks the painting context through the route
// sentinels. Registered fixture faces cover the symbol-encoded, the
// Unicode-encoded, the PUA-only (`unicode-range`) and the missing case; host
// faces (Symbol, Wingdings) are asserted only for their structural invariant.
describe.skipIf(!existsSync(chromium.executablePath()))('symbol face oracle in Chromium', () => {
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
          `export { loadNativeRouteSentinels } from ${JSON.stringify(`${root}route-sentinels.ts`)};`,
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
      const faces: [string, number[], FontFaceDescriptors][] = [
        ['Fixture Symbol Face', fonts.symbol, {}],
        ['Fixture Unicode Face', fonts.unicode, {}],
        // The symbol face admitted for its PUA codes only.
        ['Fixture PUA Face', fonts.symbol, { unicodeRange: 'U+F000-F0FF' }],
      ];
      for (const [family, bytes, descriptors] of faces) {
        const face = new FontFace(family, new Uint8Array(bytes), descriptors);
        document.fonts.add(face);
        await face.load();
      }
    }, fonts);
  }, 60_000);
  afterAll(async () => { await browser?.close(); });

  it('judges nothing in a font realm without sentinels', async () => {
    const result = await page.evaluate(() => {
      const api = (globalThis as unknown as { symbolFaceTest: Api }).symbolFaceTest;
      const ctx = document.querySelector('canvas')!.getContext('2d')!;
      const oracle = api.createSymbolFaceOracle(ctx);
      return {
        draws: oracle.draws('', '32px "Fixture Symbol Face"') ?? 'unjudged',
        // The 0.89 subset: the bullet converts, alpha does not.
        text: oracle.segments('', 'Symbol', '32px "Ooxml Absent Symbol Face 7", serif'),
      };
    });
    expect(result.draws).toBe('unjudged');
    expect(result.text).toEqual([{ text: '•', mapped: true }, { text: '', mapped: false }]);
  });

  it('asks the painting context which codes the painted font draws', async () => {
    const result = await page.evaluate(async () => {
      const api = (globalThis as unknown as { symbolFaceTest: Api }).symbolFaceTest;
      const loaded = await api.loadNativeRouteSentinels(document.fonts);
      const ctx = document.querySelector('canvas')!.getContext('2d')!;
      ctx.font = '10px sans-serif';
      ctx.direction = 'rtl';
      const oracle = api.createSymbolFaceOracle(ctx);
      const count = (font: string) => {
        let pua = 0;
        let bare = 0;
        for (let code = 0x21; code <= 0xff; code++) {
          if (oracle.draws(String.fromCharCode(0xf000 + code), font)) pua++;
          if (oracle.draws(String.fromCharCode(code), font)) bare++;
        }
        return { pua, bare };
      };
      const offscreen = api.createSymbolFaceOracle(new OffscreenCanvas(8, 8).getContext('2d'));
      return {
        loaded: loaded !== null,
        symbolFace: count('32px "Fixture Symbol Face", serif'),
        // Matching and synthesis are the engine's: a regular-only face
        // paints bold italic text too.
        synthesized: count('italic 700 32px "Fixture Symbol Face", sans-serif'),
        unicodeFace: count('32px "Fixture Unicode Face", serif'),
        // Bare codes fall through to Arial; the PUA codes are the face's.
        puaFace: count('32px "Fixture PUA Face", Arial, sans-serif'),
        // A later named family draws what the first lacks.
        later: count('32px "Ooxml Absent Symbol Face 7", "Fixture Symbol Face", serif'),
        absent: count('32px "Ooxml Absent Symbol Face 7", Arial, sans-serif'),
        generic: count('32px sans-serif'),
        hostSymbol: count('32px "Symbol", serif'),
        hostWingdings: count('32px "Wingdings", serif'),
        // A bare "(" is mirrored in this right-to-left context; the probe is not.
        paren: oracle.draws('(', '32px "Fixture Symbol Face"'),
        // DEL and SOFT HYPHEN are not in the fixture; a bare one is never native.
        softHyphen: oracle.draws('­', '32px "Fixture Symbol Face"'),
        offscreen: offscreen.draws('', '32px "Fixture Symbol Face"'),
        notACode: [oracle.draws('α', '32px "Symbol"'), oracle.draws('', '32px "Fixture Symbol Face"')],
        malformed: oracle.draws('', '"Fixture Symbol Face"') ?? 'unjudged',
        restored: [ctx.font, ctx.direction],
      };
    });
    const all = 0x7e - 0x21 + 1;
    expect(result.loaded).toBe(true);
    expect(result.symbolFace).toEqual({ pua: all, bare: all });
    expect(result.synthesized).toEqual({ pua: all, bare: all });
    // A Unicode face draws its bare codes as Latin text, never a symbol glyph.
    expect(result.unicodeFace).toEqual({ pua: 0, bare: 0 });
    expect(result.puaFace).toEqual({ pua: all, bare: 0 });
    expect(result.later).toEqual({ pua: all, bare: all });
    expect(result.absent).toEqual({ pua: 0, bare: 0 });
    expect(result.generic).toEqual({ pua: 0, bare: 0 });
    // Host faces: missing or Unicode-encoded (nothing native), or a
    // symbol-encoded face drawing most of its repertoire in both forms.
    for (const host of [result.hostSymbol, result.hostWingdings]) {
      expect(host.pua === 0 || host.pua > 180).toBe(true);
      expect(host.bare).toBeLessThanOrEqual(host.pua);
      expect(host.pua === 0 || host.bare > 180).toBe(true);
    }
    expect(result.paren).toBe(true);
    expect(result.softHyphen).toBe(false);
    expect(result.offscreen).toBe(true);
    expect(result.notACode).toEqual([false, false]);
    expect(result.malformed).toBe('unjudged');
    expect(result.restored).toEqual(['10px sans-serif', 'rtl']);
  });

  it('converts only what the painted font lacks, and never to colour', async () => {
    const result = await page.evaluate(async () => {
      const api = (globalThis as unknown as { symbolFaceTest: Api }).symbolFaceTest;
      await api.loadNativeRouteSentinels(document.fonts);
      const ctx = document.querySelector('canvas')!.getContext('2d')!;
      const oracle = api.createSymbolFaceOracle(ctx);
      const font = (family: string) => `400 32px "${family}", sans-serif`;
      const mono = api.createSymbolMonochromeProbe(ctx, '32px sans-serif');
      return {
        // 0xA0 (EURO SIGN in Symbol) is outside the fixture's repertoire.
        installed: oracle.segments('a', 'Symbol', font('Fixture Symbol Face')),
        puaOnly: oracle.segments('a', 'Symbol', font('Fixture PUA Face')),
        unicode: oracle.segments('a', 'Symbol', font('Fixture Unicode Face')),
        missing: oracle.segments('a', 'Symbol', font('Ooxml Absent Symbol Face 7')),
        missingHeart: oracle.segments('', 'Symbol', font('Ooxml Absent Symbol Face 7')),
        heartMonochrome: mono('♥︎'),
        // Wingdings 0x43 is THUMBS UP SIGN (U+1F44D): monochrome only where a
        // text face has it; otherwise the authored code is kept.
        thumbs: oracle.segments('', 'Wingdings', font('Ooxml Absent Symbol Face 7')),
        thumbsMonochrome: mono('\u{1F44D}︎'),
      };
    });
    expect(result.installed).toEqual([{ text: 'a', mapped: false }, { text: '€', mapped: true }]);
    expect(result.puaOnly).toEqual([{ text: '', mapped: false }, { text: 'α', mapped: true }]);
    expect(result.unicode).toEqual([{ text: 'αα', mapped: true }]);
    expect(result.missing).toEqual([{ text: 'αα', mapped: true }]);
    expect(result.missingHeart).toEqual(result.heartMonochrome
      ? [{ text: '♥︎', mapped: true }] : [{ text: '', mapped: false }]);
    expect(result.thumbs).toEqual(result.thumbsMonochrome
      ? [{ text: '\u{1F44D}︎', mapped: true }] : [{ text: '', mapped: false }]);
  });
});
