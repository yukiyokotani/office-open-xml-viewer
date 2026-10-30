import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { fileURLToPath } from 'node:url';

/**
 * The @font-face parser lets the host CSS parser validate font-weight,
 * font-style, font-stretch and unicode-range through probe FontFaces
 * (`packages/core/src/fonts/css-source.ts`). This checks, in Chrome, that
 * real engines mark invalid probes the way the parser assumes, in the main
 * thread and in a dedicated worker, and that probes stay invisible: no
 * request, no FontFaceSet entry, no unhandled rejection.
 */
const entry = fileURLToPath(new URL('../../packages/core/src/fonts/css-source.ts', import.meta.url));
let parserBundle: Promise<string> | undefined;
const parserScript = () => parserBundle ??= build({
  input: entry, output: { format: 'iife', name: 'cssSource' }, platform: 'browser',
}).then((bundle) => bundle.output[0].code);

const CASES: readonly (readonly [string, Record<string, string | undefined>])[] = [
  ['font-weight: 400; font-weight: calc(700);', { weight: 'calc(700)' }],
  ['font-weight: 400; font-weight: 1001;', { weight: '400' }],
  ['font-weight: 400; font-weight: heavy;', { weight: '400' }],
  ['font-weight: 300; font-weight: 100 900;', { weight: '100 900' }],
  ['font-style: normal; font-style: oblique 10deg 20deg;', { style: 'oblique 10deg 20deg' }],
  ['font-style: italic; font-style: oblique 100deg;', { style: 'italic' }],
  ['font-style: italic; font-style: slanted;', { style: 'italic' }],
  ['font-stretch: 75%; font-stretch: 50% 200%;', { stretch: '50% 200%' }],
  ['font-stretch: 75%; font-stretch: -5%;', { stretch: '75%' }],
  ['font-stretch: 75%; font-stretch: narrow;', { stretch: '75%' }],
  ['unicode-range: U+0-FF; unicode-range: U+30-20;', { unicodeRange: 'U+0-FF' }],
  ['unicode-range: U+0-FF; unicode-range: bogus;', { unicodeRange: 'U+0-FF' }],
  ['unicode-range: U+0-FF; unicode-range: U+3000-30FF, U+4E00;', { unicodeRange: 'U+3000-30FF, U+4E00' }],
];

const sheet = CASES.map(([body], index) =>
  `@font-face { font-family: Case${index}; src: url(https://fonts.example/f${index}.woff2); ${body} }`).join('\n');

/** src declarations after a valid `url(a)`; Chrome's CSSOM is the reference. */
const SRC_CASES = [
  'url(https://fonts.example/b.woff2) format("bogus")',
  'url(https://fonts.example/b.woff2) format("woff2", "woff")',
  'url(https://fonts.example/b.woff2) format(bogus), url(https://fonts.example/c.woff2)',
  'url(https://fonts.example/b.woff2) format(svg), url(https://fonts.example/c.woff2)',
  'url(https://fonts.example/b.woff2) format("embedded-opentype")',
  'url(https://fonts.example/b.woff2) format("woff2-variations")',
  'url(https://fonts.example/b.woff2) format(WOFF2)',
  'url(https://fonts.example/b.woff2) tech(bogus)',
  'url(https://fonts.example/b.woff2) tech(variations), url(https://fonts.example/c.woff2)',
  'url(https://fonts.example/b c.woff2), url(https://fonts.example/d.woff2)',
  'local(x) format(woff2)',
  'local(x)',
];

/** Stylesheets whose rule structure Chrome's CSSOM decides. */
const STRUCTURE_CASES = [
  '@media print { --x:#url({) } @font-face{font-family:Y;src:url(https://fonts.example/y.woff2)} }',
  '@media print { --x:#u } @font-face{font-family:Y;src:url(https://fonts.example/y.woff2)}',
  '@font\\-face{font-family:E;src:u\\72l(https://fonts.example/x{y.woff2)}',
  '@font-face{font-family:H;src:url(https://fonts.example/h.woff2);--x:#a(}) } @font-face{font-family:I;src:url(https://fonts.example/i.woff2)}',
];

const SHEETS = [
  ...SRC_CASES.map((src) => `@font-face{font-family:S;src:url(https://fonts.example/a.woff2);src:${src}}`),
  ...STRUCTURE_CASES,
];

interface Face { readonly family: string; readonly url: string | undefined }

/** Chrome's own reading of each sheet: top-level @font-face rules with the
 * url the library would register from the surviving src (first url entry
 * without tech() whose format is absent or registrable). */
function chromeReference(sheets: readonly string[]): Face[][] {
  const accepted = ['woff2', 'woff', 'truetype', 'opentype', 'woff2-variations', 'woff-variations',
    'truetype-variations', 'opentype-variations'];
  return sheets.map((css) => {
    const style = document.createElement('style');
    style.textContent = css;
    document.head.append(style);
    const faces: Face[] = [];
    for (const rule of Array.from((style.sheet as CSSStyleSheet).cssRules)) {
      if (!(rule instanceof CSSFontFaceRule)) continue;
      const family = rule.style.getPropertyValue('font-family').replace(/^"(.*)"$/, '$1');
      const entries = rule.style.getPropertyValue('src').split(/,\s*(?=url\(|local\()/);
      const url = entries.map((entry) => {
        const match = /^url\("([^"]*)"\)(?:\s+format\("([^"]*)"\))?$/.exec(entry.trim());
        return match && match[1] && (!match[2] || accepted.includes(match[2].toLowerCase())) ? new URL(match[1], 'https://fonts.example/kit.css').href : undefined;
      }).find((entry) => entry !== undefined);
      faces.push({ family, url });
    }
    style.remove();
    // The library drops a rule it cannot register.
    return faces.filter((entry) => entry.url !== undefined);
  });
}

interface Outcome {
  readonly host: boolean;
  readonly declarations: readonly Record<string, string | undefined>[];
  readonly fontsSize: number;
  readonly resources: number;
  readonly rejections: readonly string[];
  readonly controlStatus: string;
  readonly sheets: readonly Face[][];
}

/** Runs in either realm (`self` is the window or the worker global). */
function runInRealm(css: string, sheets: readonly string[]): Promise<Outcome> {
  const scope = self as unknown as {
    cssSource: { parseFontFaceCss(css: string, url: string): Record<string, string | undefined>[] };
    document?: { fonts: FontFaceSet };
    fonts?: FontFaceSet;
    performance: Performance;
  };
  const rejections: string[] = [];
  self.addEventListener('unhandledrejection', (event) => { rejections.push(String(event.reason)); });
  const declarations = scope.cssSource.parseFontFaceCss(css, 'https://fonts.example/kit.css');
  const parsedSheets = sheets.map((sheet) => scope.cssSource.parseFontFaceCss(sheet, 'https://fonts.example/kit.css')
    .map(({ family, url }) => ({ family: family as string, url })));
  // Control: an invalid face whose rejected `loaded` nobody observes must be
  // reported, proving the detector works and that the engine does reject it.
  void new FontFace('control-probe', 'local(control-probe)', { weight: '1001' }).loaded;
  // Give rejected probe promises time to surface as unhandled if unobserved.
  return new Promise((resolve) => setTimeout(() => resolve({
    host: typeof FontFace === 'function',
    declarations: declarations.map(({ weight, style, stretch, unicodeRange }) => ({ weight, style, stretch, unicodeRange })),
    // document.fonts on the window, self.fonts in a worker.
    fontsSize: (scope.document?.fonts ?? scope.fonts as FontFaceSet).size,
    resources: scope.performance.getEntriesByType('resource').length,
    rejections,
    controlStatus: new FontFace('control-probe', 'local(control-probe)', { weight: '1001' }).status,
    sheets: parsedSheets,
  }), 500));
}

test('host FontFace validates @font-face descriptors in the main thread and a worker', async ({ page }) => {
  test.setTimeout(60_000);
  const requests: string[] = [];
  page.context().on('request', (request) => { if (!request.url().startsWith('blob:')) requests.push(request.url()); });
  await page.goto('about:blank');
  const script = await parserScript();

  await page.addScriptTag({ content: script });
  const main = await page.evaluate(`(${runInRealm.toString()})(${JSON.stringify(sheet)}, ${JSON.stringify(SHEETS)})`) as Outcome;
  const reference = await page.evaluate(`(${chromeReference.toString()})(${JSON.stringify(SHEETS)})`) as Face[][];

  const worker = await page.evaluate(async ({ script, sheet, sheets, run }) => {
    const source = `${script}\nconst run = ${run};\nself.onmessage = async (event) => self.postMessage(await run(event.data.sheet, event.data.sheets));`;
    const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
    try {
      const w = new Worker(url);
      return await new Promise((resolve, reject) => {
        w.onmessage = (event) => resolve(event.data);
        w.onerror = (event) => reject(new Error(event.message));
        w.postMessage({ sheet, sheets });
      });
    } finally {
      URL.revokeObjectURL(url);
    }
  }, { script, sheet, sheets: SHEETS, run: runInRealm.toString() }) as Outcome;

  for (const [realm, outcome] of [['main', main], ['worker', worker]] as const) {
    expect(outcome.host, realm).toBe(true);
    expect(outcome.declarations, realm).toHaveLength(CASES.length);
    CASES.forEach(([body, expected], index) => {
      expect(outcome.declarations[index], `${realm}: ${body}`).toMatchObject(expected);
    });
    // src validity and rule structure follow Chrome's own CSSOM.
    SHEETS.forEach((css, index) => {
      expect(outcome.sheets[index], `${realm}: ${css}`).toEqual(reference[index]);
    });
    expect(outcome.fontsSize, `${realm}: probes never enter the FontFaceSet`).toBe(0);
    expect(outcome.resources, `${realm}: probes fetch nothing`).toBe(0);
    // The engine marks an invalid probe the way the parser assumes.
    expect(outcome.controlStatus, realm).toBe('error');
    // Only the deliberately unobserved control surfaces.
    expect(outcome.rejections, `${realm}: no unhandled rejection from the probes`).toHaveLength(1);
    expect(outcome.rejections[0], realm).toMatch(/SyntaxError/);
  }
  expect(requests).toEqual([]);
});

// ---- shared resolver work across sessions ------------------------------------

const sessionEntry = fileURLToPath(new URL('../../packages/core/src/internal/font-sources.ts', import.meta.url));
let sessionBundle: Promise<string> | undefined;
const sessionScript = () => sessionBundle ??= build({
  input: sessionEntry, output: { format: 'iife', name: 'fontSources' }, platform: 'browser',
}).then((bundle) => bundle.output[0].code);

interface RaceOutcome { readonly calls: number; readonly registered: number; readonly afterRelease: number }

/** Two sessions share one resolver call; the session that created it is
 * released one microtask in. The survivor must still register its face. */
async function resolverRaceInRealm(): Promise<RaceOutcome> {
  const scope = self as unknown as {
    fontSources: {
      FontSourceSession: new (options: { sources: readonly unknown[] }) => {
        ensure(demand: unknown): Promise<unknown>;
        release(): void;
      };
      fontResolver(resolve: (request: { families: readonly string[] }) => unknown[]): unknown;
    };
    document?: { fonts: FontFaceSet };
    fonts?: FontFaceSet;
  };
  const { FontSourceSession, fontResolver } = scope.fontSources;
  const set = scope.document?.fonts ?? scope.fonts as FontFaceSet;
  const before = set.size;
  let calls = 0;
  const source = fontResolver((request) => {
    calls++;
    return request.families.map((family) => ({ family, url: 'data:font/woff2;base64,AAAA' }));
  });
  const demand = { families: ['Shared'], fallbackFamilies: [], scalars: 'all' };
  const first = new FontSourceSession({ sources: [source] });
  const second = new FontSourceSession({ sources: [source] });
  const released = first.ensure(demand);
  const surviving = second.ensure(demand);
  await Promise.resolve();
  first.release();
  await released;
  await surviving;
  const registered = set.size - before;
  second.release();
  return { calls, registered, afterRelease: set.size - before };
}

test('a shared resolver call survives the release of the session that started it', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('about:blank');
  const script = await sessionScript();
  // The invalid inline face fails to load; that warning is expected.
  page.on('console', () => undefined);
  await page.addScriptTag({ content: script });
  const main = await page.evaluate(`(${resolverRaceInRealm.toString()})()`) as RaceOutcome;
  const worker = await page.evaluate(async ({ script, run }) => {
    const source = `${script}\nconst run = ${run};\nself.onmessage = async () => self.postMessage(await run());`;
    const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
    try {
      const w = new Worker(url);
      return await new Promise((resolve, reject) => {
        w.onmessage = (event) => resolve(event.data);
        w.onerror = (event) => reject(new Error(event.message));
        w.postMessage(null);
      });
    } finally {
      URL.revokeObjectURL(url);
    }
  }, { script, run: resolverRaceInRealm.toString() }) as RaceOutcome;
  for (const [realm, outcome] of [['main', main], ['worker', worker]] as const) {
    expect(outcome, realm).toEqual({ calls: 1, registered: 1, afterRelease: 0 });
  }
});
