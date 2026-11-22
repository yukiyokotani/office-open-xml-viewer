import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { build } from 'vite';
import { existsSync } from 'node:fs';
import { symbolFixtureFontBytes } from '../testing/symbol-font-fixture';

type Api = Pick<typeof import('./native-route-line-metrics'), 'createNamedFamilyOracle'>
  & Pick<typeof import('./route-sentinels'), 'loadNativeRouteSentinels'>;

interface RealmCapability {
  /** Before this realm registered the sentinels. */
  before: boolean;
  sentinelsApply: boolean;
  available: boolean;
  lang: boolean;
}

// Real Chromium: the capability a run language rests on is judged in the
// realm that measures. Each realm (the page, an OffscreenCanvas of the page,
// an iframe document, a worker) applies the route sentinels only once they
// are registered in its own FontFaceSet, and the oracle reports exactly that.
describe.skipIf(!existsSync(chromium.executablePath()))('named-family oracle in Chromium', () => {
  let browser: Browser;
  let page: Page;
  let bundle = '';
  beforeAll(async () => {
    const root = new URL('./', import.meta.url).pathname;
    const entry = '/__named-family-oracle-test.ts';
    const result = await build({
      configFile: false, logLevel: 'error',
      plugins: [{ name: 'named-family-oracle-test', resolveId: id => id === entry ? id : null,
        load: id => id === entry ? [
          `export { createNamedFamilyOracle } from ${JSON.stringify(`${root}native-route-line-metrics.ts`)};`,
          `export { loadNativeRouteSentinels } from ${JSON.stringify(`${root}route-sentinels.ts`)};`,
        ].join('\n') : null }],
      build: { write: false, minify: false, lib: { entry, formats: ['iife'], name: 'namedFamilyTest' } },
    });
    const outputs = Array.isArray(result) ? result : [result];
    const chunk = outputs.flatMap(output => 'output' in output ? output.output : [])
      .find(output => output.type === 'chunk');
    if (!chunk || chunk.type !== 'chunk') throw new Error('bundle missing');
    bundle = chunk.code;
    browser = await chromium.launch({ channel: 'chromium' });
    page = await browser.newPage();
    await page.setContent('<!doctype html><canvas></canvas>');
    await page.addScriptTag({ content: bundle });
  }, 120_000);
  afterAll(async () => { await browser?.close(); }, 60_000);

  it('applies the sentinels only in the realm that registered them', async () => {
    const result = await page.evaluate(async (code) => {
      const api = (globalThis as unknown as { namedFamilyTest: Api }).namedFamilyTest;
      const capability = (context: unknown, before: boolean): RealmCapability => {
        const oracle = api.createNamedFamilyOracle(context);
        return { before, sentinelsApply: oracle.sentinelsApply, available: oracle.available,
          lang: 'lang' in (context as object) };
      };
      const ctx = document.querySelector('canvas')!.getContext('2d')!;
      ctx.font = '10px sans-serif';
      const pageBefore = api.createNamedFamilyOracle(ctx).sentinelsApply;

      // A worker registers in its own set; the page's set is still empty.
      // A worker registers in its own set. As in a layout session, the
      // sentinel families are first resolved after their registration: a
      // Chromium worker keeps its first resolution of a family name, so this
      // realm is not probed beforehand.
      const source = `${code};self.onmessage = async () => {
        const api = self.namedFamilyTest;
        const loaded = await api.loadNativeRouteSentinels(self.fonts);
        const context = new OffscreenCanvas(1, 1).getContext('2d');
        const oracle = api.createNamedFamilyOracle(context);
        self.postMessage({ loaded: loaded !== null, before: false, sentinelsApply: oracle.sentinelsApply,
          available: oracle.available, lang: 'lang' in context });
      };`;
      const worker = new Worker(URL.createObjectURL(new Blob([source], { type: 'text/javascript' })));
      const inWorker = await new Promise<RealmCapability & { loaded: boolean }>((resolve, reject) => {
        worker.onmessage = (event) => resolve(event.data);
        worker.onerror = (event) => reject(new Error(event.message));
        worker.postMessage(0);
      });
      worker.terminate();
      const pageAfterWorker = api.createNamedFamilyOracle(ctx).sentinelsApply;

      const loaded = await api.loadNativeRouteSentinels(document.fonts);
      const inPage = capability(ctx, pageBefore);
      const offscreen = capability(new OffscreenCanvas(1, 1).getContext('2d'), false);

      // An iframe canvas is shaped by the iframe document's own set.
      const frame = document.createElement('iframe');
      document.body.append(frame);
      const frameDocument = frame.contentDocument!;
      const frameCanvas = frameDocument.createElement('canvas');
      frameDocument.body.append(frameCanvas);
      const frameContext = frameCanvas.getContext('2d')!;
      const frameBefore = api.createNamedFamilyOracle(frameContext).sentinelsApply;
      const frameLoaded = await api.loadNativeRouteSentinels(frameDocument.fonts);
      const inFrame = capability(frameContext, frameBefore);
      return {
        loaded: loaded !== null, frameLoaded: frameLoaded !== null, pageAfterWorker,
        inPage, offscreen, inFrame, inWorker, fontAfter: ctx.font,
      };
    }, bundle);

    expect(result.loaded).toBe(true);
    expect(result.inWorker.loaded).toBe(true);
    // Unregistered: nothing is applied, and another realm's registration
    // does not reach this one.
    expect([result.inPage.before, result.pageAfterWorker, result.inFrame.before]).toEqual([false, false, false]);
    // Registered: the measuring context and both proof contexts apply them.
    expect([result.inPage.sentinelsApply, result.offscreen.sentinelsApply, result.inWorker.sentinelsApply])
      .toEqual([true, true, true]);
    // An iframe realm judges for itself: the page's registration did not
    // reach it, its own does.
    expect([result.frameLoaded, result.inFrame.sentinelsApply]).toEqual([true, true]);
    // A run language additionally needs Canvas `lang`, the same in every realm.
    for (const realm of [result.inPage, result.offscreen, result.inWorker, result.inFrame]) {
      expect(realm.available).toBe(realm.sentinelsApply && realm.lang);
    }
    expect(new Set([result.inPage.lang, result.offscreen.lang, result.inWorker.lang]).size).toBe(1);
    // The measuring context keeps its own font.
    expect(result.fontAfter).toBe('10px sans-serif');
  }, 120_000);

  it('admits only text a named family of the route draws', async () => {
    const font = Array.from(symbolFixtureFontBytes({ cmap: 'unicode', first: 0x21, last: 0x7e, name: 'FixtureUni' }));
    const result = await page.evaluate(async (bytes) => {
      const api = (globalThis as unknown as { namedFamilyTest: Api }).namedFamilyTest;
      const face = new FontFace('Fixture Unicode Face', new Uint8Array(bytes));
      document.fonts.add(face);
      await face.load();
      await api.loadNativeRouteSentinels(document.fonts);
      const ctx = document.querySelector('canvas')!.getContext('2d')!;
      const oracle = api.createNamedFamilyOracle(ctx);
      const draws = (list: string, text: string) => oracle.draws(list, 400, 'normal', text);
      return {
        available: oracle.available,
        covered: draws('"Fixture Unicode Face", serif', 'abc!~'),
        // The fixture has no Cyrillic: the generic would draw it.
        fallback: draws('"Fixture Unicode Face", serif', 'aж'),
        // A generic before the named family is the evidence boundary.
        genericFirst: draws('serif, "Fixture Unicode Face"', 'abc'),
        // An absent family before the covering one is skipped by matching.
        absentFirst: draws('"Ooxml Absent Face 7", "Fixture Unicode Face", serif', 'abc'),
        absentOnly: draws('"Ooxml Absent Face 7", serif', 'abc'),
        // Presentation-selected clusters are outside the evidence boundary.
        emoji: draws('"Fixture Unicode Face", serif', '\u{1f469}'),
      };
    }, font);
    // Without Canvas `lang` nothing is admitted; with it, exactly the text
    // named families draw.
    expect(result).toEqual({
      available: result.available,
      covered: result.available,
      fallback: false,
      genericFirst: false,
      absentFirst: result.available,
      absentOnly: false,
      emoji: false,
    });
  }, 120_000);
});
